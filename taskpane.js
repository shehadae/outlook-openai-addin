async function summarizeEmail() {
  const item = Office.context.mailbox.item;
  const body = await new Promise(resolve => item.body.getAsync(Office.CoercionType.Text, result => resolve(result.value)));
  const apiKey = 'YOUR_OPENAI_API_KEY_HERE';

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'gpt-4o',
      messages: [{ role: 'user', content: `Summarize this email: ${body}` }],
      temperature: 0.3
    })
  });

  const data = await response.json();
  const summary = data.choices?.[0]?.message?.content;
  document.getElementById('output').innerText = summary || 'No summary generated.';
}
