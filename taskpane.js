// Initialize Office Add-in
Office.onReady((result) => {
  if (result.host === Office.HostType.Outlook) {
    // Set up event listeners for buttons
    document.getElementById('analyze-btn').onclick = analyzeEmail;
    document.getElementById('generate-btn').onclick = generateReply;
  }
});

// Analyze the current email
function analyzeEmail() {
  const statusEl = document.getElementById('status');
  statusEl.textContent = 'Analyzing email...';
  
  try {
    // Get current email item from Outlook
    const item = Office.context.mailbox.item;
    const subject = item.subject;
    const body = item.body.text || '';
    
    console.log('Analyzing email: ' + subject);
    
    // Sentiment Analysis
    const sentiment = analyzeSentiment(body);
    document.getElementById('sentiment-value').textContent = sentiment;
    
    // Urgency Detection
    const urgency = detectUrgency(body, subject);
    document.getElementById('urgency-value').textContent = urgency;
    
    // Extract Key Points
    const keyPoints = extractKeyPoints(body);
    const list = document.getElementById('key-points-list');
    list.innerHTML = '';
    keyPoints.forEach(point => {
      const li = document.createElement('li');
      li.textContent = point;
      list.appendChild(li);
    });
    
    statusEl.textContent = 'Email analysis complete!';
  } catch (error) {
    statusEl.textContent = 'Error: ' + error.message;
    console.error('Analysis error:', error);
  }
}

// Generate AI Reply
function generateReply() {
  const statusEl = document.getElementById('status');
  statusEl.textContent = 'Generating reply...';
  
  try {
    const item = Office.context.mailbox.item;
    const body = item.body.text || '';
    const reply = generateSmartReply(body);
    
    // Display generated reply
    const replyEl = document.getElementById('email-analysis');
    if (replyEl) {
      replyEl.value = reply;
    }
    
    statusEl.textContent = 'Reply generated. Copy and paste into your response.';
  } catch (error) {
    statusEl.textContent = 'Error: ' + error.message;
    console.error('Generation error:', error);
  }
}

// Analyze sentiment of text
function analyzeSentiment(text) {
  if (!text) return 'Neutral';
  
  const positive = (text.match(/great|excellent|amazing|wonderful|outstanding|good|perfect|fantastic/gi) || []).length;
  const negative = (text.match(/bad|terrible|awful|horrible|disappointed|poor|hate|terrible|worst/gi) || []).length;
  
  if (positive > negative) return 'Positive';
  if (negative > positive) return 'Negative';
  return 'Neutral';
}

// Detect urgency level
function detectUrgency(body, subject) {
  const text = (body + ' ' + subject).toLowerCase();
  const urgentTerms = ['urgent', 'asap', 'immediately', 'emergency', 'critical', 'deadline', 'today', 'now'];
  
  const urgentCount = urgentTerms.filter(term => text.includes(term)).length;
  
  if (urgentCount >= 2) return 'High';
  if (urgentCount === 1) return 'Medium';
  return 'Low';
}

// Extract key points from text
function extractKeyPoints(text) {
  if (!text) return [];
  
  const sentences = text.split(/[.!?]+/).filter(s => s.trim().length > 0);
  return sentences
    .filter(s => s.trim().length > 20)
    .slice(0, 3)
    .map(s => s.trim());
}

// Generate smart reply based on sentiment
function generateSmartReply(emailBody) {
  const sentiment = analyzeSentiment(emailBody);
  let greeting = 'Hi,';
  let response = '';
  
  if (sentiment === 'Positive') {
    response = 'Thank you for your positive message. I appreciate your feedback and will respond to your points shortly.';
  } else if (sentiment === 'Negative') {
    response = 'I understand your concerns and sincerely appreciate you bringing them to my attention. Let me address these issues promptly.';
  } else {
    response = 'Thank you for reaching out. I have reviewed your message and will provide a detailed response.';
  }
  
  return greeting + '\n\n' + response + '\n\nBest regards';
}
