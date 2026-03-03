import { useState } from 'react';
import { FileLoader } from './components/FileLoader';
import { RoadmapGrid } from './components/RoadmapGrid';
import { HeatmapGrid } from './components/HeatmapGrid';
import { parseRoadmap } from './utils/excel';
import type { RoadmapData } from './utils/excel';
import './styles.css';

const SOURCE_URL =
  'https://mckessoncorp.sharepoint.com/:x:/s/GRPProductCommercialLeadershipTeam/IQB5wgqUs0k5SLOkUoR3V4JNAf9kLnoYyclcwZUNPoqpNiE?e=t0jlnX';

function App() {
  const [roadmap, setRoadmap] = useState<RoadmapData | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [view, setView] = useState<'detail' | 'heatmap'>('detail');

  const handleLoad = (buffer: ArrayBuffer) => {
    setIsLoading(true);
    setError(null);
    setRoadmap(null);

    try {
      const data = parseRoadmap(buffer);
      setRoadmap(data);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Unknown error occurred';
      setError(errorMessage);
      console.error('Excel parsing error:', err);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="app">
      <header className="app-header">
        <h1>Priority Roadmap</h1>
        <p>Product initiatives by team and quarter</p>
      </header>

      <main className="app-main">
        <FileLoader onLoad={handleLoad} isLoading={isLoading} error={error} sourceUrl={SOURCE_URL} />
        {roadmap && (
          <>
            <div className="view-toggle">
              <button
                className={`view-toggle-btn${view === 'detail' ? ' view-toggle-btn--active' : ''}`}
                onClick={() => setView('detail')}
              >
                Detail View
              </button>
              <button
                className={`view-toggle-btn${view === 'heatmap' ? ' view-toggle-btn--active' : ''}`}
                onClick={() => setView('heatmap')}
              >
                Heatmap View
              </button>
            </div>
            {view === 'detail' ? (
              <RoadmapGrid data={roadmap} />
            ) : (
              <HeatmapGrid data={roadmap} />
            )}
          </>
        )}
      </main>
    </div>
  );
}

export default App;
