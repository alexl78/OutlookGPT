import { Client } from 'openai';

Office.initialize = function (reason) {
    document.getElementById('ask-btn')?.addEventListener('click', askChatGPT);
    document.getElementById('cancel-btn')?.addEventListener('click', closeModal);
    document.getElementById('query-input')?.addEventListener('input', handleInput);

    // Handlers for saving API key and model
    document.getElementById('save-api-key-btn')?.addEventListener('click', saveApiKey);
    document.getElementById('save-model-btn')?.addEventListener('click', saveModel);

    // Load saved settings upon initialization
    loadApiKey();
    loadModel();
};

// Global variable to store the previous query
let previousQuery = '';

function launchGPTModal() {
    const modal = document.querySelector('.modal') as HTMLElement;
    modal.style.display = 'flex';
    const input = document.getElementById('query-input') as HTMLTextAreaElement;
    input.value = previousQuery;
}

function closeModal() {
    const modal = document.querySelector('.modal') as HTMLElement;
    modal.style.display = 'none';
}

function handleInput() {
    const query = (document.getElementById('query-input') as HTMLTextAreaElement).value;
    const askBtn = document.getElementById('ask-btn') as HTMLButtonElement;
    askBtn.disabled = query.trim() === '';
}

async function askChatGPT() {
    const query = (document.getElementById('query-input') as HTMLTextAreaElement).value;
    previousQuery = query;

    try {
        // Get the current email content
        const item = Office.context.mailbox.item;
        const emailContent = await new Promise<string>((resolve, reject) => {
            item.body.getAsync(Office.CoercionType.Html, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject('Failed to get email content');
                }
            });
        });

        // Get the API key and model
        const apiKey = Office.context.roamingSettings.get('openaiApiKey');
        const model = Office.context.roamingSettings.get('openaiModel') || 'GPT-4o';  // Default value

        if (!apiKey) {
            throw new Error('API key is not set. Please configure your API key in settings.');
        }

        // Send the query to ChatGPT
        const response = await sendToChatGPT(query, emailContent, apiKey, model);

        // Insert the response as a reply
        item.displayReplyForm(response);
    } catch (error) {
        console.error('Error interacting with ChatGPT or Outlook:', error);
    }

    closeModal();
}

// Function to send a query to ChatGPT using the selected model
async function sendToChatGPT(query: string, emailContent: string, apiKey: string, model: string) {
    const openai = new Client({
        apiKey: apiKey
    });

    const prompt = `You are responding to the following email: ${emailContent}\n\nUser's query: ${query}`;
    const completion = await openai.createCompletion({
        model: model,
        prompt: prompt,
        max_tokens: 1500,
        temperature: 0.7
    });

    return completion.data.choices[0].text;
}

// Function to save the API key in roamingSettings
function saveApiKey() {
    const apiKey = (document.getElementById('api-key-input') as HTMLInputElement).value;
    if (apiKey) {
        Office.context.roamingSettings.set('openaiApiKey', apiKey);
        Office.context.roamingSettings.saveAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                alert('API Key saved successfully.');
