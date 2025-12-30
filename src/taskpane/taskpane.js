Office.onReady(function(reason) {
    if (reason === Office.Initialize Reason.ContentInitialized) {
        document.getElementById('analyze-btn').onclick = analyzeEmail;
        document.getElementById('generate-btn').onclick = generateReply;
    }
});

function analyzeEmail() {
    Excel.run(async function(context) {
        try {
            const body = context.mailbox.item.body;
            const subject = context.mailbox.item.subject;
            
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
            
            document.getElementById('status').textContent = 'Email analysis complete!';
        } catch (error) {
            console.log('Error: ' + JSON.stringify(error));
        }
    });
}

function generateReply() {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = result.value;
            const reply = generateSmartReply(emailBody);
            
            document.getElementById('email-analysis').value = reply;
            document.getElementById('status').textContent = 'Reply generated. Copy and paste into your response.';
        }
    });
}

function analyzeSentiment(text) {
    const positive = (text.match(/great|excellent|amazing|wonderful|outstanding/gi) || []).length;
    const negative = (text.match(/bad|terrible|awful|horrible|disappointed/gi) || []).length;
    
    if (positive > negative) return 'Positive';
    if (negative > positive) return 'Negative';
    return 'Neutral';
}

function detectUrgency(body, subject) {
    const urgentTerms = ['urgent', 'asap', 'immediately', 'emergency', 'critical', 'deadline'];
    const urgentCount = urgentTerms.filter(term => 
        body.toLowerCase().includes(term) || subject.toLowerCase().includes(term)
    ).length;
    
    if (urgentCount >= 2) return 'High';
    if (urgentCount === 1) return 'Medium';
    return 'Low';
}

function extractKeyPoints(text) {
    const sentences = text.split(/[.!?]+/);
    return sentences
        .filter(s => s.trim().length > 20)
        .slice(0, 3)
        .map(s => s.trim());
}

function generateSmartReply(emailBody) {
    const sentiment = analyzeSentiment(emailBody);
    let greeting = 'Hi,';
    let response = '';
    
    if (sentiment === 'Positive') {
        response = 'Thank you for your positive message. I appreciate your feedback.';
    } else if (sentiment === 'Negative') {
        response = 'I understand your concerns. Let me help address these issues.';
    } else {
        response = 'Thank you for reaching out. I am here to assist you.';
    }
    
    return greeting + '\n\n' + response + '\n\nBest regards';
}
