
import fs from 'fs';
const content = fs.readFileSync('dist/index.html', 'utf8');
const scriptMatch = content.match(/<script\b[^>]*>([\s\S]*?)<\/script>/i);
if (scriptMatch) {
  const script = scriptMatch[1];
  console.log("Script length:", script.length);
  console.log("First 500 chars:", script.substring(0, 500));
  console.log("Last 500 chars:", script.substring(script.length - 500));
  
  // Check for specific problematic keywords
  const keywords = ['import.meta', 'URL(', 'fetch(', 'import('];
  keywords.forEach(kw => {
    if (script.includes(kw)) {
      console.log(`Contains "${kw}": Yes`);
    } else {
      console.log(`Contains "${kw}": No`);
    }
  });
} else {
  console.log("No script tag found.");
}
