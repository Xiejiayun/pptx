#!/usr/bin/env node
import { runCli } from '@pptx/cli';

process.exitCode = await runCli(process.argv.slice(2));
