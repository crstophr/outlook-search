/**
 * LLM Integration
 * 
 * Provides intelligent analysis and relevance scoring for email search results.
 * Connects to OpenClaw's model provider for natural language understanding.
 */

/**
 * Default model configuration - matches OpenClaw setup
 */
const DEFAULT_CONFIG = {
    baseUrl: 'http://localhost:8080/v1',
    apiKey: 'llamacpp-local',
    model: 'Qwen3.5-27B-Claude-4.6-Opus-Reasoning-Distilled.Q4_K_M.gguf'
};

/**
 * Analyze a batch of email summaries and score them for relevance
 */
async function analyzeEmailRelevance(emails, searchContext, options = {}) {
    const { config = DEFAULT_CONFIG } = options;
    
    if (emails.length === 0) {
        return { scored: [], filtered: [] };
    }
    
    // Build the email context string
    const emailText = emails.map((email, idx) => `
Email #${idx + 1}:
- Subject: ${email.subject || '(no subject)'}
- From: ${email.senderName || ''} <${email.senderEmail || ''}>\n- Date: ${email.receivedDateTime ? new Date(email.receivedDateTime).toLocaleString() : 'unknown'}
- Has attachments: ${email.hasAttachments}
`).join('\n');
    
    // Build the analysis prompt
    const prompt = buildAnalysisPrompt(searchContext, emailText);
    
    try {
        const response = await callLlm(prompt, {
            system: 'You are an intelligent email search assistant. Analyze emails and score their relevance to what the user is looking for.',
            config
        });
        
        // Parse the LLM's JSON response
        const analysis = parseAnalysisResponse(response);
        
        // Score the emails based on analysis
        const scored = emails.map((email, idx) => ({
            ...email,
            relevanceScore: analysis[idx]?.score || 5,
            reasoning: analysis[idx]?.reasoning || '',
            hasDocuments: analysis[idx]?.hasDocuments || 'unknown',
            searchMatches: analysis[idx]?.searchMatches || []
        }));
        
        // Filter and sort by relevance
        const filtered = scored
            .filter(e => e.relevanceScore >= 4)  // Minimum threshold of 4
            .sort((a, b) => b.relevanceScore - a.relevanceScore);
        
        return { scored, filtered };
    } catch (error) {
        console.error('LLM analysis failed:', error.message);
        return fallbackAnalysis(emails, searchContext);
    }
}

/**
 * Build the prompt for analyzing email relevance
 */
function buildAnalysisPrompt(searchContext, emailText) {
    const { query, context = '' } = searchContext;
    
    return `
The user is searching for something in their Outlook emails.

**What they're looking for:** "${query}"

${context ? `**Additional context they provided:** ${context}` : ''}

**Search results (first 50 emails by date/subject):**

${emailText}

---

**YOUR TASK:** Analyze each email and determine how relevant it is to what the user wants.

For EACH email, provide:

1. **SCORE** (1-10):
   - 10: Almost certainly what they want
   - 8-9: Very likely relevant
   - 6-7: Probably relevant
   - 4-5: Maybe relevant, worth checking
   - 1-3: Unlikely to be relevant

2. **REASONING** (one sentence): Why you gave this score

3. **HAS_DOCUMENTS** (yes/no/maybe): Does it seem to mention attachments or documents?

4. **SEARCH_MATCHES**: Which of these did it match? (keywords, sender pattern, date range, document indicators)

---

**OUTPUT FORMAT:** Return ONLY a JSON array, like this:

${JSON.stringify([
    {
        "index": 1,
        "score": 9,
        "reasoning": "Subject directly mentions both 'Visa' and 'statement of work'",
        "hasDocuments": "yes",
        "searchMatches": ["keywords", "document indicators"]
    },
    {
        "index": 2,
        "score": 3,
        "reasoning": "Only mentions Visa, no reference to statement of work or documents",
        "hasDocuments": "no",
        "searchMatches": ["keywords"]
    }
], null, 2)}

Do not include any text outside the JSON array.
`;
}

/**
 * Call the LLM with a prompt
 */
async function callLlm(prompt, options = {}) {
    const { system, config = DEFAULT_CONFIG } = options;
    
    const messages = [];
    
    if (system) {
        messages.push({ role: 'system', content: system });
    }
    
    messages.push({ role: 'user', content: prompt });
    
    const response = await fetch(`${config.baseUrl}/chat/completions`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${config.apiKey}`
        },
        body: JSON.stringify({
            model: config.model,
            messages: messages,
            temperature: 0.3,  // Low temperature for consistent analysis
            max_tokens: 4096
        })
    });
    
    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`LLM API error: ${response.status} - ${errorText}`);
    }
    
    const data = await response.json();
    return data.choices[0]?.message?.content || '';
}

/**
 * Parse the LLM's analysis response into structured data
 */
function parseAnalysisResponse(response) {
    // Extract JSON from response (might have markdown code blocks)
    let jsonText = response.trim();
    
    // Remove markdown code fences if present
    const fenceMatch = jsonText.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (fenceMatch) {
        jsonText = fenceMatch[1].trim();
    }
    
    try {
        return JSON.parse(jsonText);
    } catch (error) {
        console.error('Failed to parse LLM response as JSON:', error.message);
        console.log('Response was:', response.substring(0, 500));
        return [];
    }
}

/**
 * Fallback analysis when LLM is unavailable
 */
function fallbackAnalysis(emails, searchContext) {
    const { query } = searchContext;
    const queryLower = query.toLowerCase();
    
    console.log('  Using fallback keyword-based analysis...');
    
    const scored = emails.map(email => {
        let score = 5; // Neutral baseline
        const subject = (email.subject || '').toLowerCase();
        const sender = (email.senderEmail || '').toLowerCase();
        
        // Check for direct query matches in subject
        const queryWords = queryLower.split(/\s+/).filter(w => w.length > 2);
        const matchedWords = queryWords.filter(word => 
            subject.includes(word) || sender.includes(word)
        );
        
        if (matchedWords.length > 0) {
            score += Math.min(3, matchedWords.length);
        }
        
        // Boost for document indicators
        const docIndicators = ['attached', 'attachment', 'pdf', 'document', 'file'];
        if (docIndicators.some(ind => subject.includes(ind))) {
            score += 2;
        }
        
        // Boost for hasAttachments flag
        if (email.hasAttachments) {
            score += 1;
        }
        
        return {
            ...email,
            relevanceScore: Math.min(10, score),
            reasoning: `Matched ${matchedWords.length} query words${email.hasAttachments ? ', has attachments' : ''}`,
            hasDocuments: email.hasAttachments ? 'yes' : 'maybe',
            searchMatches: matchedWords.length > 0 ? ['keywords'] : []
        };
    });
    
    const filtered = scored
        .filter(e => e.relevanceScore >= 4)
        .sort((a, b) => b.relevanceScore - a.relevanceScore);
    
    return { scored, filtered };
}

/**
 * Extract relevant snippets from email body based on search context
 */
async function extractRelevantSnippets(emailBody, searchContext, options = {}) {
    const { query } = searchContext;
    const { config = DEFAULT_CONFIG } = options;
    
    if (!emailBody || emailBody.length === 0) {
        return [];
    }
    
    // Truncate body if too long
    const maxBodyLength = 8000;
    const truncatedBody = emailBody.length > maxBodyLength 
        ? emailBody.substring(0, maxBodyLength) + '...[truncated]'
        : emailBody;
    
    const prompt = `
The user is searching for: "${query}"

Extract the MOST relevant parts of this email body that relate to their search.

Email body:
${truncatedBody}

Return a JSON array with up to 3 snippets:
[
  {
    "text": "The exact excerpt from the email",
    "relevance": "Why this is relevant (keywords matched, mentions documents, etc.)"
  }
]
`;
    
    try {
        const response = await callLlm(prompt, { config });
        return parseSnippetResponse(response);
    } catch (error) {
        console.error('Snippet extraction failed:', error.message);
        return fallbackSnippetExtraction(truncatedBody, query);
    }
}

/**
 * Parse snippet extraction response
 */
function parseSnippetResponse(response) {
    let jsonText = response.trim();
    const fenceMatch = jsonText.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (fenceMatch) {
        jsonText = fenceMatch[1].trim();
    }
    
    try {
        return JSON.parse(jsonText);
    } catch (error) {
        console.error('Failed to parse snippet response');
        return [];
    }
}

/**
 * Fallback snippet extraction using keyword matching
 */
function fallbackSnippetExtraction(body, query) {
    const queryWords = query.toLowerCase().split(/\s+/).filter(w => w.length > 3);
    const sentences = body.split(/[.!?]\s*/);
    
    const snippets = sentences.slice(0, 50).map(sentence => ({
        text: sentence.trim() + '.',
        relevance: 'Keyword match'
    })).filter(snippet => 
        queryWords.some(word => snippet.text.toLowerCase().includes(word))
    ).slice(0, 3);
    
    return snippets;
}

module.exports = {
    analyzeEmailRelevance,
    extractRelevantSnippets,
    callLlm,
    DEFAULT_CONFIG
};
