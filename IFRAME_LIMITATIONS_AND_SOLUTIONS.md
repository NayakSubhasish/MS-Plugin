# X-Frame-Options and Office Add-in Integration Solutions

## The Problem: X-Frame-Options: SAMEORIGIN

Many websites send the `X-Frame-Options: SAMEORIGIN` response header, which prevents them from being displayed in iframes unless they're hosted on the same domain as the parent page. This is a security feature designed to prevent:

- **Clickjacking attacks** - Malicious sites overlaying invisible iframes over legitimate content
- **UI redressing** - Tricking users into clicking on elements they think are from one site but are actually from another
- **Cross-site request forgery (CSRF)** - Unauthorized actions performed on behalf of the user

## Why This Affects Office Add-ins

Office Add-ins run in a web-based taskpane that operates under the same security restrictions as regular web browsers. When you try to embed external websites using iframes, you'll encounter this limitation.

## Solutions for Office Add-ins

### 1. **API-Based Integration** (Recommended)

Instead of embedding websites, use their APIs to fetch data and display it natively in your add-in.

#### Example: Weather Service Integration
```javascript
// Instead of: <iframe src="https://weather.com/widget"></iframe>

const fetchWeatherData = async (city) => {
  const response = await fetch(`https://api.openweathermap.org/data/2.5/weather?q=${city}&appid=YOUR_API_KEY`);
  const data = await response.json();
  
  // Display data in your own UI components
  return {
    temperature: data.main.temp,
    description: data.weather[0].description,
    icon: data.weather[0].icon
  };
};
```

#### Example: AI Service Integration (Like Your Current Gemini Setup)
```javascript
// Your current approach is perfect - using Gemini API directly
const getSuggestedReply = async (prompt) => {
  const response = await fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${API_KEY}`
    },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }]
    })
  });
  
  return response.json();
};
```

### 2. **Server-Side Proxy**

Create a backend service that fetches content from external sites and serves it to your add-in.

#### Backend Proxy Example (Node.js/Express)
```javascript
// server.js
const express = require('express');
const axios = require('axios');
const app = express();

app.post('/api/proxy', async (req, res) => {
  try {
    const { url } = req.body;
    const response = await axios.get(url);
    
    // Process the content (remove scripts, sanitize HTML, etc.)
    const sanitizedContent = sanitizeHtml(response.data);
    
    res.json({ content: sanitizedContent });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

function sanitizeHtml(html) {
  // Remove potentially dangerous elements
  return html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, '')
    .replace(/on\w+="[^"]*"/gi, '');
}
```

#### Frontend Usage
```javascript
const fetchExternalContent = async (url) => {
  const response = await fetch('/api/proxy', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ url })
  });
  
  const data = await response.json();
  return data.content;
};
```

### 3. **Webhook Integration**

Send data to external services and receive responses via webhooks.

```javascript
const sendToExternalService = async (data) => {
  const response = await fetch('https://external-service.com/webhook', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      event: 'email_processed',
      data: data,
      timestamp: new Date().toISOString()
    })
  });
  
  return response.json();
};
```

### 4. **Content Embedding Alternatives**

Instead of iframes, embed rich content using native components:

#### Charts and Visualizations
```javascript
// Use charting libraries like Chart.js, D3.js, or Recharts
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip } from 'recharts';

const ChartComponent = ({ data }) => (
  <LineChart width={400} height={200} data={data}>
    <CartesianGrid strokeDasharray="3 3" />
    <XAxis dataKey="name" />
    <YAxis />
    <Tooltip />
    <Line type="monotone" dataKey="value" stroke="#8884d8" />
  </LineChart>
);
```

#### Forms and Interactive Elements
```javascript
// Create native form components instead of embedding external forms
const ContactForm = () => (
  <form onSubmit={handleSubmit}>
    <input type="text" placeholder="Name" />
    <input type="email" placeholder="Email" />
    <textarea placeholder="Message" />
    <button type="submit">Send</button>
  </form>
);
```

### 5. **Deep Linking**

Instead of embedding, provide links that open external content in new tabs/windows.

```javascript
const ExternalLink = ({ url, children }) => (
  <Link href={url} target="_blank" rel="noopener noreferrer">
    {children}
  </Link>
);

// Usage
<ExternalLink url="https://external-service.com/dashboard">
  View Dashboard
</ExternalLink>
```

## Security Considerations

### CORS (Cross-Origin Resource Sharing)
- Ensure APIs you're calling support CORS
- Use your own domain in the `Origin` header
- Consider using a proxy for non-CORS-compliant APIs

### API Key Management
- Never expose API keys in client-side code
- Use environment variables and server-side proxy
- Implement proper authentication and authorization

### Content Sanitization
- Always sanitize HTML content from external sources
- Remove potentially dangerous scripts and iframes
- Use libraries like DOMPurify for sanitization

## Best Practices for Office Add-ins

1. **Prefer APIs over iframes** - APIs provide better control and security
2. **Implement proper error handling** - Handle network failures gracefully
3. **Cache responses** - Improve performance by caching API responses
4. **Use loading states** - Provide user feedback during API calls
5. **Respect rate limits** - Implement throttling for external API calls
6. **Test thoroughly** - Ensure your add-in works in different Office environments

## Example Implementation in Your Project

Your current Mail Magic add-in already follows best practices by using the Gemini API directly. To extend this pattern:

```javascript
// Add more AI services
const services = {
  gemini: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent',
  openai: 'https://api.openai.com/v1/chat/completions',
  claude: 'https://api.anthropic.com/v1/messages'
};

const callAIService = async (service, prompt) => {
  const config = {
    gemini: {
      url: services.gemini,
      headers: { 'Authorization': `Bearer ${GEMINI_API_KEY}` },
      body: { contents: [{ parts: [{ text: prompt }] }] }
    },
    openai: {
      url: services.openai,
      headers: { 'Authorization': `Bearer ${OPENAI_API_KEY}` },
      body: { model: 'gpt-3.5-turbo', messages: [{ role: 'user', content: prompt }] }
    }
  };
  
  const serviceConfig = config[service];
  const response = await fetch(serviceConfig.url, {
    method: 'POST',
    headers: { ...serviceConfig.headers, 'Content-Type': 'application/json' },
    body: JSON.stringify(serviceConfig.body)
  });
  
  return response.json();
};
```

## Conclusion

While `X-Frame-Options: SAMEORIGIN` limits iframe usage, it actually encourages better architectural patterns. By using APIs, webhooks, and native components, you can create more secure, performant, and maintainable Office Add-ins that provide rich functionality without compromising security. 