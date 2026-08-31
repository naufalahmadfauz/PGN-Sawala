import type { PgnValidationIssue } from "./pgn-types";

export interface ParsedCellTurn {
  turnNumber: number;
  userInput: string;
}

export interface ParsedCellTurns {
  turns: ParsedCellTurn[];
  issues: PgnValidationIssue[];
}

function stripWrappingQuotes(value: string): string {
  const trimmed = value.trim();
  const pairs: Array<[string, string]> = [
    ['"', '"'],
    ["'", "'"],
    ["“", "”"],
    ["‘", "’"],
  ];
  for (const [opening, closing] of pairs) {
    if (trimmed.startsWith(opening) && trimmed.endsWith(closing)) {
      return trimmed.slice(opening.length, -closing.length).trim();
    }
  }
  return trimmed;
}

export function parseTurnsFromCell(
  value: string,
  sheetName: string,
  rowNumber: number,
): ParsedCellTurns {
  const input = value.replace(/\r\n/g, "\n").trim();
  const markerPattern = /(?:^|\n)\s*Turn\s+(\d+)\s*:\s*/gi;
  const matches = [...input.matchAll(markerPattern)];
  if (matches.length === 0) {
    return {
      turns: input ? [{ turnNumber: 1, userInput: input }] : [],
      issues: [],
    };
  }

  const issues: PgnValidationIssue[] = [];
  const turns: ParsedCellTurn[] = [];
  const prefix = input.slice(0, matches[0].index).trim();
  if (prefix) {
    issues.push({
      code: "INVALID_TURN",
      severity: "ERROR",
      sheetName,
      rowNumber,
      message: "Text appears before the first explicit Turn marker.",
    });
  }

  for (let index = 0; index < matches.length; index += 1) {
    const match = matches[index];
    const turnNumber = Number(match[1]);
    const start = (match.index ?? 0) + match[0].length;
    const end = matches[index + 1]?.index ?? input.length;
    const userInput = stripWrappingQuotes(input.slice(start, end));
    if (!userInput) {
      issues.push({
        code: "MISSING_USER_INPUT",
        severity: "ERROR",
        sheetName,
        rowNumber,
        message: `Turn ${turnNumber} has no User Input.`,
      });
    }
    turns.push({ turnNumber, userInput });
  }

  const seenTurns = new Set<number>();
  for (let index = 0; index < turns.length; index += 1) {
    const turn = turns[index];
    const expected = index + 1;
    if (
      !Number.isInteger(turn.turnNumber) ||
      turn.turnNumber < 1 ||
      seenTurns.has(turn.turnNumber) ||
      turn.turnNumber !== expected
    ) {
      issues.push({
        code: "INVALID_TURN",
        severity: "ERROR",
        sheetName,
        rowNumber,
        message: `Explicit turns must be unique and sequential from Turn 1; found Turn ${turn.turnNumber} at position ${expected}.`,
      });
    }
    seenTurns.add(turn.turnNumber);
  }

  return { turns, issues };
}
