import * as XLSX from 'xlsx';

// SheetJS returns dynamic cell values — using a union here to avoid `any`
// while still covering all possible Excel cell types.
type CellValue = string | number | boolean | Date | null | undefined;

export interface Initiative {
  name: string;
  suite: string;
  priority: string;
  quarter: string;
  value: string;
  uspOwner: string;
  primaryDevTeam: string;
  team: string;
  workType: string;
  strategyPillar: string;
  dependencies: string;
}

export interface TeamInfo {
  name: string;
  suite: string;
}

export interface RoadmapData {
  teams: TeamInfo[];
  quarters: string[];
  pillars: string[];
  grid: Map<string, Map<string, Initiative[]>>;
}

export async function fetchExcelFile(url: string): Promise<ArrayBuffer> {
  const response = await fetch(url, {
    method: 'GET',
    headers: {
      'Accept': '*/*',
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch file: ${response.status} ${response.statusText}`);
  }

  return response.arrayBuffer();
}

const SHEET_NAME = 'Product Initiatives';
const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];

// Column indices in the "Product Initiatives" sheet
const COL = {
  INITIATIVE: 0,  // A
  SUITE: 1,        // B
  PRIORITY: 2,     // C
  QUARTER: 3,      // D
  VALUE: 4,        // E
  USP_OWNER: 5,    // F
  PRIMARY_DEV: 6,  // G
  SINGLE_TEAM: 7,  // H
  WORK_TYPE: 8,        // I
  DEPENDENCIES: 9,     // J
  STRATEGY_PILLAR: 10, // K (to be added)
} as const;

function str(val: CellValue): string {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

export function parseRoadmap(arrayBuffer: ArrayBuffer): RoadmapData {
  const workbook = XLSX.read(arrayBuffer, {
    type: 'array',
    cellFormula: true,
    cellNF: true,
  });

  const ws = workbook.Sheets[SHEET_NAME];
  if (!ws) {
    const available = workbook.SheetNames.join(', ');
    throw new Error(`Sheet "${SHEET_NAME}" not found. Available sheets: ${available}`);
  }

  const rows = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    blankrows: false,
    defval: '',
  }) as CellValue[][];

  // Find Strategy Pillar column dynamically from the header row
  const headerRow = rows[0] || [];
  const pillarColIdx = headerRow.findIndex(
    (cell) => str(cell).toLowerCase() === 'strategy pillar',
  );

  // Skip header row (index 0)
  const dataRows = rows.slice(1);

  // Build the grid: team -> quarter -> initiatives[]
  const grid = new Map<string, Map<string, Initiative[]>>();
  const teamSuiteMap = new Map<string, string>();
  const pillarSet = new Set<string>();

  for (const row of dataRows) {
    const team = str(row[COL.SINGLE_TEAM]);
    if (!team) continue; // skip rows with no team

    const quarter = str(row[COL.QUARTER]).toUpperCase();
    if (!QUARTERS.includes(quarter)) continue; // skip rows without a valid quarter

    const suite = str(row[COL.SUITE]);

    const initiative: Initiative = {
      name: str(row[COL.INITIATIVE]),
      suite,
      priority: str(row[COL.PRIORITY]),
      quarter,
      value: str(row[COL.VALUE]),
      uspOwner: str(row[COL.USP_OWNER]),
      primaryDevTeam: str(row[COL.PRIMARY_DEV]),
      team,
      workType: str(row[COL.WORK_TYPE]),
      strategyPillar: pillarColIdx >= 0 ? str(row[pillarColIdx]) : '',
      dependencies: str(row[COL.DEPENDENCIES]),
    };

    const pillar = pillarColIdx >= 0 ? str(row[pillarColIdx]) : '';
    if (pillar) pillarSet.add(pillar);

    if (!teamSuiteMap.has(team)) {
      teamSuiteMap.set(team, suite);
    }

    if (!grid.has(team)) {
      grid.set(team, new Map());
    }
    const teamMap = grid.get(team)!;

    if (!teamMap.has(quarter)) {
      teamMap.set(quarter, []);
    }
    teamMap.get(quarter)!.push(initiative);
  }

  // Sort by suite first, then team name
  const teams: TeamInfo[] = Array.from(teamSuiteMap.entries())
    .map(([name, suite]) => ({ name, suite }))
    .sort((a, b) => a.suite.localeCompare(b.suite) || a.name.localeCompare(b.name));

  const pillars = Array.from(pillarSet).sort();

  return { teams, quarters: QUARTERS, pillars, grid };
}

export function validateUrl(url: string): { valid: boolean; error?: string; warning?: string } {
  if (!url.trim()) {
    return { valid: false, error: 'Please enter a URL' };
  }

  try {
    const urlObj = new URL(url);
    if (!urlObj.protocol.startsWith('http')) {
      return { valid: false, error: 'URL must start with http or https' };
    }
  } catch {
    return { valid: false, error: 'Invalid URL format' };
  }

  const lowerUrl = url.toLowerCase();
  const looksLikeExcel =
    lowerUrl.endsWith('.xlsx') || lowerUrl.endsWith('.xls') ||
    lowerUrl.includes('.xlsx?') || lowerUrl.includes('.xls?') ||
    lowerUrl.includes('sharepoint') || lowerUrl.includes('onedrive') ||
    lowerUrl.includes('download=1');

  if (!looksLikeExcel) {
    return {
      valid: true,
      warning: 'This URL doesn\'t look like an Excel or SharePoint link. It may still work if it serves a downloadable file.',
    };
  }

  return { valid: true };
}
