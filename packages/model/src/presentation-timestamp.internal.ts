import { type CoreQualifiedTypeDescriptor } from './presentation-core-properties.internal.js';

export const W3CDTF_QUALIFIED_TYPE: CoreQualifiedTypeDescriptor = {
  attributeNamespace: 'http://www.w3.org/2001/XMLSchema-instance',
  attributePreferredPrefix: 'xsi',
  valueNamespace: 'http://purl.org/dc/terms/',
  valuePreferredPrefix: 'dcterms',
  valueLocalName: 'W3CDTF',
};

const PRESENTATION_TIMESTAMP_PATTERN =
  /^([0-9]{4})-([0-9]{2})-([0-9]{2})T([0-9]{2}):([0-9]{2}):([0-9]{2})(?:\.([0-9]+))?(Z|[+-][0-9]{2}:[0-9]{2})$/;

export function normalizePresentationTimestamp(
  value: unknown,
  propertyName: 'createdAt' | 'modifiedAt',
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string' || !isPresentationTimestamp(value)) {
    throw new TypeError(
      `Presentation ${propertyName} must be a valid W3CDTF date-time string or undefined`,
    );
  }
  return value;
}

export function isPresentationTimestamp(value: string): boolean {
  const match = PRESENTATION_TIMESTAMP_PATTERN.exec(value);
  if (!match) return false;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);
  const timezone = match[8]!;
  if (year < 1 || month < 1 || month > 12) return false;
  const leapYear = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const daysInMonth = [31, leapYear ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (day < 1 || day > daysInMonth[month - 1]!) return false;
  if (hour > 23 || minute > 59 || second > 59) return false;
  if (timezone !== 'Z') {
    const offsetHour = Number(timezone.slice(1, 3));
    const offsetMinute = Number(timezone.slice(4, 6));
    if (
      offsetHour > 14
      || offsetMinute > 59
      || (offsetHour === 14 && offsetMinute !== 0)
    ) return false;
  }
  return true;
}
