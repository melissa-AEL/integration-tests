import { useEffect, useState } from 'react';
import { app, authentication } from '@microsoft/teams-js';

function App() {
  const [status, setStatus] = useState('Initializing...');
  const [token, setToken] = useState(null);
  const [sites, setSites] = useState(null);
  const [error, setError] = useState(null);

  useEffect(() => {
    const init = async () => {
      try {
        setStatus('Initializing Teams SDK...');
        await app.initialize();

        setStatus('Getting auth token...');
        authentication.getAuthToken({
          resources: ['https://graph.microsoft.com'],
          successCallback: async (token) => {
            setToken(token);
            setStatus('Token acquired. Fetching SharePoint sites...');

            const response = await fetch('https://graph.microsoft.com/v1.0/sites?search=*', {
              headers: {
                Authorization: `Bearer ${token}`
              }
            });

            if (!response.ok) {
              throw new Error(`Graph API error: ${response.status}`);
            }

            const data = await response.json();
            setSites(data.value);
            setStatus('Sites loaded successfully.');
          },
          failureCallback: (err) => {
            setError(`Token error: ${err}`);
            setStatus('Failed to get token.');
          }
        });
      } catch (err) {
        setError(`Init error: ${err.message}`);
        setStatus('Initialization failed.');
      }
    };

    init();
  }, []);

  return (
    <div style={{ fontFamily: 'sans-serif', padding: '1rem' }}>
      <h2>Teams + SharePoint Integration Test</h2>
      <p><strong>Status:</strong> {status}</p>

      {error && <p style={{ color: 'red' }}><strong>Error:</strong> {error}</p>}

      {token && (
        <details>
          <summary>Access Token (first 100 chars)</summary>
          <code>{token.substring(0, 100)}...</code>
        </details>
      )}

      {sites && (
        <div>
          <h3>SharePoint Sites</h3>
          <ul>
            {sites.map(site => (
              <li key={site.id}>{site.name} ({site.webUrl})</li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}

export default App;