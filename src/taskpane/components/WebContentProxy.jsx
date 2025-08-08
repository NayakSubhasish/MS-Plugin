import * as React from "react";
import { Button, makeStyles, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
  container: {
    padding: "16px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: "16px",
  },
  content: {
    maxHeight: "300px",
    overflowY: "auto",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusSmall,
    marginTop: "8px",
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    padding: "8px",
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderRadius: tokens.borderRadiusSmall,
  }
});

const WebContentProxy = () => {
  const styles = useStyles();
  const [content, setContent] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState("");

  // Example: Fetch content through your own backend proxy
  const fetchExternalContent = async (url) => {
    setLoading(true);
    setError("");
    
    try {
      // This would call your backend proxy service
      const response = await fetch('/api/proxy', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ url })
      });
      
      if (!response.ok) {
        throw new Error('Failed to fetch content');
      }
      
      const data = await response.json();
      setContent(data.content);
    } catch (err) {
      setError(`Error fetching content: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <h3>External Content Integration</h3>
      <p>Instead of iframes, use APIs or server-side proxies to integrate external content.</p>
      
      <Button 
        onClick={() => fetchExternalContent('https://example.com')}
        disabled={loading}
        appearance="secondary"
      >
        {loading ? "Loading..." : "Fetch External Content"}
      </Button>
      
      {error && <div className={styles.error}>{error}</div>}
      
      {content && (
        <div className={styles.content}>
          <h4>Fetched Content:</h4>
          <div dangerouslySetInnerHTML={{ __html: content }} />
        </div>
      )}
    </div>
  );
};

export default WebContentProxy; 