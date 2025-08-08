import * as React from "react";
import { Button, makeStyles, tokens, Text, Link } from "@fluentui/react-components";

const useStyles = makeStyles({
  container: {
    padding: "16px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: "16px",
  },
  section: {
    marginBottom: "16px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusSmall,
  },
  buttonGroup: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
    marginTop: "8px",
  },
  content: {
    maxHeight: "200px",
    overflowY: "auto",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusSmall,
    marginTop: "8px",
    fontSize: tokens.fontSizeBase300,
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    padding: "8px",
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderRadius: tokens.borderRadiusSmall,
    marginTop: "8px",
  },
  success: {
    color: tokens.colorPaletteGreenForeground1,
    padding: "8px",
    backgroundColor: tokens.colorPaletteGreenBackground1,
    borderRadius: tokens.borderRadiusSmall,
    marginTop: "8px",
  }
});

const ExternalServiceIntegration = () => {
  const styles = useStyles();
  const [results, setResults] = React.useState({});
  const [loading, setLoading] = React.useState({});

  // 1. Direct API Integration
  const callExternalAPI = async (service) => {
    setLoading(prev => ({ ...prev, [service]: true }));
    
    try {
      let response;
      switch (service) {
        case 'weather':
          // Example: Weather API
          response = await fetch('https://api.openweathermap.org/data/2.5/weather?q=London&appid=YOUR_API_KEY');
          break;
        case 'news':
          // Example: News API
          response = await fetch('https://newsapi.org/v2/top-headlines?country=us&apiKey=YOUR_API_KEY');
          break;
        case 'translate':
          // Example: Translation API
          response = await fetch('https://translation.googleapis.com/language/translate/v2?key=YOUR_API_KEY', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              q: 'Hello world',
              target: 'es'
            })
          });
          break;
        default:
          throw new Error('Unknown service');
      }
      
      const data = await response.json();
      setResults(prev => ({ ...prev, [service]: data }));
    } catch (error) {
      setResults(prev => ({ ...prev, [service]: { error: error.message } }));
    } finally {
      setLoading(prev => ({ ...prev, [service]: false }));
    }
  };

  // 2. Webhook Integration
  const triggerWebhook = async () => {
    setLoading(prev => ({ ...prev, webhook: true }));
    
    try {
      const response = await fetch('https://your-webhook-endpoint.com/webhook', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          event: 'email_processed',
          data: { timestamp: new Date().toISOString() }
        })
      });
      
      const result = await response.json();
      setResults(prev => ({ ...prev, webhook: result }));
    } catch (error) {
      setResults(prev => ({ ...prev, webhook: { error: error.message } }));
    } finally {
      setLoading(prev => ({ ...prev, webhook: false }));
    }
  };

  // 3. Content Embedding (without iframe)
  const embedContent = (type) => {
    const content = {
      'chart': {
        type: 'chart',
        data: { labels: ['Jan', 'Feb', 'Mar'], values: [10, 20, 15] },
        config: { type: 'bar' }
      },
      'calendar': {
        type: 'calendar',
        events: [
          { date: '2024-01-15', title: 'Meeting', description: 'Team sync' },
          { date: '2024-01-20', title: 'Deadline', description: 'Project due' }
        ]
      },
      'form': {
        type: 'form',
        fields: [
          { name: 'name', label: 'Name', type: 'text' },
          { name: 'email', label: 'Email', type: 'email' }
        ]
      }
    };
    
    setResults(prev => ({ ...prev, [type]: content[type] }));
  };

  return (
    <div className={styles.container}>
      <Text size={500} weight="semibold">External Service Integration Examples</Text>
      <Text size={200} style={{ marginBottom: "16px" }}>
        These examples show how to integrate external services without using iframes.
      </Text>

      {/* API Integration Section */}
      <div className={styles.section}>
        <Text size={400} weight="semibold">1. Direct API Integration</Text>
        <Text size={200}>Call external APIs directly from your add-in</Text>
        <div className={styles.buttonGroup}>
          <Button 
            onClick={() => callExternalAPI('weather')}
            disabled={loading.weather}
            appearance="secondary"
            size="small"
          >
            {loading.weather ? "Loading..." : "Weather API"}
          </Button>
          <Button 
            onClick={() => callExternalAPI('news')}
            disabled={loading.news}
            appearance="secondary"
            size="small"
          >
            {loading.news ? "Loading..." : "News API"}
          </Button>
          <Button 
            onClick={() => callExternalAPI('translate')}
            disabled={loading.translate}
            appearance="secondary"
            size="small"
          >
            {loading.translate ? "Loading..." : "Translation API"}
          </Button>
        </div>
        
        {results.weather && (
          <div className={styles.content}>
            <Text size={200} weight="semibold">Weather API Response:</Text>
            <pre>{JSON.stringify(results.weather, null, 2)}</pre>
          </div>
        )}
      </div>

      {/* Webhook Section */}
      <div className={styles.section}>
        <Text size={400} weight="semibold">2. Webhook Integration</Text>
        <Text size={200}>Send data to external services via webhooks</Text>
        <Button 
          onClick={triggerWebhook}
          disabled={loading.webhook}
          appearance="secondary"
          size="small"
        >
          {loading.webhook ? "Sending..." : "Trigger Webhook"}
        </Button>
        
        {results.webhook && (
          <div className={styles.content}>
            <Text size={200} weight="semibold">Webhook Response:</Text>
            <pre>{JSON.stringify(results.webhook, null, 2)}</pre>
          </div>
        )}
      </div>

      {/* Content Embedding Section */}
      <div className={styles.section}>
        <Text size={400} weight="semibold">3. Content Embedding</Text>
        <Text size={200}>Embed rich content without iframes</Text>
        <div className={styles.buttonGroup}>
          <Button 
            onClick={() => embedContent('chart')}
            appearance="secondary"
            size="small"
          >
            Embed Chart
          </Button>
          <Button 
            onClick={() => embedContent('calendar')}
            appearance="secondary"
            size="small"
          >
            Embed Calendar
          </Button>
          <Button 
            onClick={() => embedContent('form')}
            appearance="secondary"
            size="small"
          >
            Embed Form
          </Button>
        </div>
        
        {results.chart && (
          <div className={styles.content}>
            <Text size={200} weight="semibold">Chart Data:</Text>
            <pre>{JSON.stringify(results.chart, null, 2)}</pre>
          </div>
        )}
      </div>

      {/* Best Practices */}
      <div className={styles.section}>
        <Text size={400} weight="semibold">Best Practices</Text>
        <ul style={{ margin: "8px 0", paddingLeft: "20px" }}>
          <li>Use APIs instead of iframes when possible</li>
          <li>Implement proper error handling and loading states</li>
          <li>Cache responses to improve performance</li>
          <li>Use CORS-compliant APIs or proxy through your backend</li>
          <li>Consider rate limiting and API quotas</li>
        </ul>
      </div>
    </div>
  );
};

export default ExternalServiceIntegration; 