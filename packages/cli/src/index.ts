import { promises as fs } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { Command, CommanderError, InvalidArgumentError } from 'commander';
import { OpcPackage } from '@pptx/opc';
import { PptxDocument, type CompatibilityProfile } from '@pptx/sdk';
import { diffPackages, inspectPackage } from '@pptx/testkit';

export interface CliIo {
  readonly stdout: (text: string) => void;
  readonly stderr: (text: string) => void;
}

const defaultIo: CliIo = {
  stdout: (text) => process.stdout.write(text),
  stderr: (text) => process.stderr.write(text),
};

export async function runCli(argv: readonly string[], io: CliIo = defaultIo): Promise<number> {
  const program = createProgram(io);
  try {
    await program.parseAsync(['node', 'pptx-inspect', ...argv]);
    return 0;
  } catch (error) {
    if (error instanceof CommanderError && ['commander.helpDisplayed', 'commander.version'].includes(error.code)) return 0;
    const message = error instanceof Error ? error.message : String(error);
    if (program.opts<{ json?: boolean }>().json) {
      io.stdout(`${JSON.stringify({ ok: false, error: { code: errorCode(error), message } })}\n`);
    } else {
      io.stderr(`Error: ${message}\n`);
    }
    return error instanceof CommanderError ? error.exitCode || 2 : 1;
  }
}

function createProgram(io: CliIo): Command {
  const program = new Command();
  program
    .name('pptx-inspect')
    .description('Inspect, validate, diff, and make narrow edits to PPTX OOXML packages')
    .version('0.1.0')
    .option('--json', 'emit stable machine-readable JSON')
    .exitOverride()
    .configureOutput({ writeOut: io.stdout, writeErr: io.stderr });

  program
    .command('doctor')
    .description('check the local runtime and optional rendering dependency')
    .action(() => {
      const nodeMajor = Number(process.versions.node.split('.')[0]);
      const soffice = commandExists('soffice');
      emit(program, io, 'doctor', {
        version: '0.1.0',
        node: { version: process.versions.node, supported: nodeMajor >= 20 },
        platform: process.platform,
        auth: { required: false, source: 'not-required' },
        mode: 'offline',
        optional: { libreoffice: soffice },
      });
    });

  const packageCommand = program.command('package').description('read package-level structure and diagnostics');
  packageCommand
    .command('inspect <file>')
    .description('list parts, content types, hashes, and relationship counts')
    .action(async (file: string) => emit(program, io, 'package.inspect', await inspectPackage(file)));
  packageCommand
    .command('validate <file>')
    .description('run OPC, relationship, codec, and compatibility validation')
    .option('--profile <name>', 'compatibility profile', 'powerpoint-current')
    .action(async (file: string, options: { profile: CompatibilityProfile }) => {
      const document = await PptxDocument.open(file);
      await document.write({ compatibility: options.profile, mode: 'permissive' });
      const errors = document.diagnostics.filter(({ severity }) => severity === 'error');
      emit(program, io, 'package.validate', {
        valid: errors.length === 0,
        errorCount: errors.length,
        warningCount: document.diagnostics.filter(({ severity }) => severity === 'warning').length,
        diagnostics: document.diagnostics,
      });
    });
  packageCommand
    .command('diff <before> <after>')
    .description('compare decompressed part payload hashes')
    .action(async (before: string, after: string) => emit(program, io, 'package.diff', await diffPackages(before, after)));

  const slides = program.command('slides').description('discover slides and make narrow slide edits');
  slides
    .command('list <file>')
    .description('list slide numbers, part URIs, titles, and shape counts')
    .action(async (file: string) => {
      const document = await PptxDocument.open(file);
      emit(
        program,
        io,
        'slides.list',
        document.slides.map((slide, index) => ({
          number: index + 1,
          partUri: slide.partUri,
          title: slide.title.text,
          shapeCount: slide.shapes.length,
        })),
      );
    });
  slides
    .command('set-title <file>')
    .description('set one slide title; requires --out unless --dry-run is used')
    .requiredOption('--slide <number>', 'one-based slide number', positiveInteger)
    .requiredOption('--text <text>', 'new title text')
    .option('--out <file>', 'output PPTX path')
    .option('--dry-run', 'validate the requested edit without writing')
    .action(
      async (
        file: string,
        options: { slide: number; text: string; out?: string; dryRun?: boolean },
      ) => {
        if (!options.dryRun && !options.out) throw new Error('--out is required unless --dry-run is used');
        const document = await PptxDocument.open(file);
        const slide = document.slides[options.slide - 1];
        if (!slide) throw new RangeError(`Slide ${options.slide} is out of range`);
        const before = slide.title.text;
        slide.title.text = options.text;
        await document.write({ mode: 'strict' });
        if (!options.dryRun) await document.writeFile(options.out!);
        emit(program, io, 'slides.set-title', {
          dryRun: Boolean(options.dryRun),
          slide: options.slide,
          before,
          after: options.text,
          output: options.dryRun ? null : options.out,
        });
      },
    );

  const part = program.command('part').description('raw read-only access to an exact OPC part');
  part
    .command('read <file> <uri>')
    .description('read one exact part; binary JSON output is base64')
    .option('--out <file>', 'write raw part bytes to a file')
    .action(async (file: string, uri: string, options: { out?: string }) => {
      const pkg = await OpcPackage.open(new Uint8Array(await fs.readFile(file)));
      const value = pkg.requirePart(uri);
      if (options.out) await fs.writeFile(options.out, value.bytes);
      const textual = /xml|text|json/.test(value.contentType);
      emit(program, io, 'part.read', {
        uri: value.uri,
        contentType: value.contentType,
        bytes: value.bytes.byteLength,
        output: options.out ?? null,
        encoding: options.out ? 'file' : textual ? 'utf8' : 'base64',
        content: options.out
          ? null
          : textual
            ? new TextDecoder().decode(value.bytes)
            : Buffer.from(value.bytes).toString('base64'),
      });
    });

  return program;
}

function emit(program: Command, io: CliIo, command: string, data: unknown): void {
  if (program.opts<{ json?: boolean }>().json) {
    io.stdout(`${JSON.stringify({ ok: true, command, data })}\n`);
    return;
  }
  if (typeof data === 'string') io.stdout(`${data}\n`);
  else io.stdout(`${JSON.stringify(data, null, 2)}\n`);
}

function positiveInteger(value: string): number {
  const number = Number(value);
  if (!Number.isInteger(number) || number < 1) throw new InvalidArgumentError('expected a positive integer');
  return number;
}

function commandExists(command: string): boolean {
  return spawnSync('sh', ['-c', `command -v ${command}`], { stdio: 'ignore' }).status === 0;
}

function errorCode(error: unknown): string {
  if (error instanceof CommanderError) return 'CLI_USAGE_ERROR';
  if (error instanceof RangeError) return 'CLI_RANGE_ERROR';
  return 'CLI_ERROR';
}

