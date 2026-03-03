import { useMemo, useState } from 'react';
import type { RoadmapData } from '../utils/excel';

// 8 distinguishable hues derived from "Refreshing Summer Fun" palette
const PILLAR_PALETTE = [
  { border: '#219EBC', bg: 'rgba(33, 158, 188, 0.08)' },    // blue green
  { border: '#FB8500', bg: 'rgba(251, 133, 0, 0.08)' },     // tiger orange
  { border: '#023047', bg: 'rgba(2, 48, 71, 0.10)' },       // deep space blue
  { border: '#06B6D4', bg: 'rgba(6, 182, 212, 0.08)' },      // cyan
  { border: '#D63384', bg: 'rgba(214, 51, 132, 0.08)' },     // warm rose (complement)
  { border: '#10B981', bg: 'rgba(16, 185, 129, 0.08)' },     // emerald
  { border: '#2D6A4F', bg: 'rgba(45, 106, 79, 0.08)' },     // forest green (complement)
  { border: '#9B5DE5', bg: 'rgba(155, 93, 229, 0.08)' },    // purple (complement)
];

const NONE_STYLE = { border: '#CED4DA', bg: '#ffffff' };

interface RoadmapGridProps {
  data: RoadmapData;
}

export const RoadmapGrid: React.FC<RoadmapGridProps> = ({ data }) => {
  const { teams, quarters, pillars, grid } = data;
  const [expanded, setExpanded] = useState<string | null>(null);
  const [suiteFilter, setSuiteFilter] = useState<string>('');
  const [pillarFilter, setPillarFilter] = useState<string>('');

  const suites = useMemo(
    () => Array.from(new Set(teams.map((t) => t.suite))).sort(),
    [teams],
  );

  // Map each pillar value to a color from the palette
  const pillarColors = useMemo(() => {
    const map = new Map<string, { border: string; bg: string }>();
    pillars.forEach((p, i) => {
      map.set(p, PILLAR_PALETTE[i % PILLAR_PALETTE.length]);
    });
    return map;
  }, [pillars]);

  const filteredTeams = suiteFilter
    ? teams.filter((t) => t.suite === suiteFilter)
    : teams;

  const toggleCell = (team: string, quarter: string) => {
    const key = `${team}|${quarter}`;
    setExpanded((prev) => (prev === key ? null : key));
  };

  const getPillarStyle = (pillar: string) => {
    return pillarColors.get(pillar) || NONE_STYLE;
  };

  return (
    <div>
      <div className="roadmap-filters">
        <label className="filter-label" htmlFor="suite-filter">Suite:</label>
        <select
          id="suite-filter"
          className="filter-select"
          value={suiteFilter}
          onChange={(e) => setSuiteFilter(e.target.value)}
        >
          <option value="">All</option>
          {suites.map((s) => (
            <option key={s} value={s}>{s}</option>
          ))}
        </select>

        {pillars.length > 0 && (
          <div className="pillar-legend">
            {pillars.map((p) => {
              const color = getPillarStyle(p);
              const isActive = pillarFilter === p;
              return (
                <button
                  key={p}
                  className={`pillar-legend-item pillar-legend-item--clickable${isActive ? ' pillar-legend-item--active' : ''}`}
                  onClick={() => {
                    if (isActive) {
                      setPillarFilter('');
                    } else {
                      setPillarFilter(p);
                      setSuiteFilter('');
                    }
                  }}
                >
                  <span
                    className="pillar-legend-swatch"
                    style={{ backgroundColor: color.border }}
                  />
                  {p}
                </button>
              );
            })}
          </div>
        )}
      </div>

      <div className="roadmap-grid-wrapper">
        <table className="roadmap-grid">
          <thead>
            <tr>
              <th className="roadmap-team-header">Team</th>
              <th className="roadmap-suite-header">Suite</th>
              {quarters.map((q) => (
                <th key={q} className="roadmap-quarter-header">{q}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredTeams.map((team) => {
              const teamMap = grid.get(team.name);
              return (
                <tr key={team.name}>
                  <td className="roadmap-team-cell">{team.name}</td>
                  <td className="roadmap-suite-cell">{team.suite}</td>
                  {quarters.map((q) => {
                    const allInitiatives = teamMap?.get(q) || [];
                    const initiatives = pillarFilter
                      ? allInitiatives.filter((init) => init.strategyPillar === pillarFilter)
                      : allInitiatives;
                    const cellKey = `${team.name}|${q}`;
                    const isExpanded = expanded === cellKey;

                    return (
                      <td
                        key={q}
                        className={`roadmap-cell${initiatives.length > 0 ? ' roadmap-cell--has-items' : ''}${isExpanded ? ' roadmap-cell--expanded' : ''}`}
                        onClick={() => initiatives.length > 0 && toggleCell(team.name, q)}
                      >
                        {initiatives.map((init, i) => {
                          const color = getPillarStyle(init.strategyPillar);
                          return (
                            <div
                              key={i}
                              className="roadmap-initiative"
                              style={{
                                borderLeftColor: color.border,
                                backgroundColor: color.bg,
                              }}
                            >
                              <div className="roadmap-initiative-name">{init.name}</div>
                              {isExpanded && init.value && (
                                <div className="roadmap-initiative-value">{init.value}</div>
                              )}
                            </div>
                          );
                        })}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};
