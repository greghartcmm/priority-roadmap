import { useMemo } from 'react';
import type { RoadmapData } from '../utils/excel';

const PILLAR_PALETTE = [
  '#219EBC', // blue green
  '#FB8500', // tiger orange
  '#023047', // deep space blue
  '#06B6D4', // cyan
  '#D63384', // warm rose
  '#10B981', // emerald
  '#2D6A4F', // forest green
  '#9B5DE5', // purple
];

const NONE_COLOR = '#CED4DA';

// Choose white or dark text based on background luminance
function textColor(hex: string): string {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return lum > 0.55 ? '#023047' : '#ffffff';
}

interface PillarPill {
  pillar: string;
  color: string;
  count: number;
}

interface HeatmapGridProps {
  data: RoadmapData;
}

export const HeatmapGrid: React.FC<HeatmapGridProps> = ({ data }) => {
  const { teams, quarters, pillars, grid } = data;

  const pillarColorMap = useMemo(() => {
    const map = new Map<string, string>();
    pillars.forEach((p, i) => {
      map.set(p, PILLAR_PALETTE[i % PILLAR_PALETTE.length]);
    });
    return map;
  }, [pillars]);

  // Group teams by suite for section headers
  const teamsBySuite = useMemo(() => {
    const groups: { suite: string; teams: typeof teams }[] = [];
    let currentSuite = '';
    let currentGroup: typeof teams = [];

    for (const team of teams) {
      if (team.suite !== currentSuite) {
        if (currentGroup.length > 0) {
          groups.push({ suite: currentSuite, teams: currentGroup });
        }
        currentSuite = team.suite;
        currentGroup = [];
      }
      currentGroup.push(team);
    }
    if (currentGroup.length > 0) {
      groups.push({ suite: currentSuite, teams: currentGroup });
    }
    return groups;
  }, [teams]);

  // Group initiatives by unique pillar, return one pill per pillar
  const getPills = (teamName: string, quarter: string): PillarPill[] => {
    const teamMap = grid.get(teamName);
    const initiatives = teamMap?.get(quarter) || [];
    if (initiatives.length === 0) return [];

    const counts = new Map<string, number>();
    for (const init of initiatives) {
      const key = init.strategyPillar || '';
      counts.set(key, (counts.get(key) || 0) + 1);
    }

    // Sort by pillar order for consistency
    return Array.from(counts.entries())
      .sort(([a], [b]) => {
        if (!a) return 1;
        if (!b) return -1;
        return pillars.indexOf(a) - pillars.indexOf(b);
      })
      .map(([pillar, count]) => ({
        pillar: pillar || 'None',
        color: pillar ? (pillarColorMap.get(pillar) || NONE_COLOR) : NONE_COLOR,
        count,
      }));
  };

  return (
    <div className="heatmap-container">
      <div className="heatmap-grid-wrapper">
        <table className="heatmap-grid">
          <thead>
            <tr>
              <th className="heatmap-suite-header">Suite</th>
              <th className="heatmap-team-header">Team</th>
              {quarters.map((q) => (
                <th key={q} className="heatmap-quarter-header">{q}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {teamsBySuite.map((group) =>
              group.teams.map((team, teamIdx) => (
                <tr key={team.name} className="heatmap-team-row">
                  {teamIdx === 0 && (
                    <td
                      className="heatmap-suite-cell"
                      rowSpan={group.teams.length}
                    >
                      {group.suite}
                    </td>
                  )}
                  <td className="heatmap-team-cell">{team.name}</td>
                  {quarters.map((q) => {
                    const pills = getPills(team.name, q);
                    return (
                      <td key={q} className="heatmap-data-cell">
                        {pills.length > 0 && (
                          <div className="heatmap-pills">
                            {pills.map((pill, i) => (
                              <span
                                key={i}
                                className="heatmap-pill"
                                style={{
                                  backgroundColor: pill.color,
                                  color: textColor(pill.color),
                                }}
                                title={`${pill.pillar} (${pill.count} initiative${pill.count > 1 ? 's' : ''})`}
                              >
                                {pill.pillar}
                              </span>
                            ))}
                          </div>
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Legend */}
      <div className="heatmap-legend">
        {pillars.map((p) => (
          <span key={p} className="heatmap-legend-item">
            <span
              className="heatmap-legend-swatch"
              style={{ backgroundColor: pillarColorMap.get(p) }}
            />
            {p}
          </span>
        ))}
      </div>
    </div>
  );
};
