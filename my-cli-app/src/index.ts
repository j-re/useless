import * as fs from 'fs';
import * as readline from 'readline';
import fetch from 'node-fetch';

// Parse continueFromId from command line arguments
const args = process.argv.slice(2);
let continueFromId: number | undefined = undefined;
for (const arg of args) {
    console.log('Args:', arg);
    if (arg.startsWith('--continueFromId=')) {
        const val = arg.split('=')[1];
        if (val) continueFromId = Number(val);
    }
}

console.log('Continue from ID:', continueFromId);

const inputCsv = 'input.csv'; // Update with your CSV filename
const logFile = 'api_responses.log';

type TestApiData = {
    data: {
        username: string;
        // Add other properties if needed
        id: number;
        status: string;
    };
    success: boolean;
};

type LogEntry = {
    name: string;
    id: string;
    username: string;
    testApiResponse: TestApiData;
    suspendApiResponse: any; // You can define a more specific type if you know the structure
};

async function processCsv() {
    const fileStream = fs.createReadStream(inputCsv);
    const rl = readline.createInterface({
        input: fileStream,
        crlfDelay: Infinity
    });

    // Skip header
    let isFirstLine = true;
    let skipping = continueFromId !== undefined;

    for await (const line of rl) {
        if (isFirstLine) {
            isFirstLine = false;
            continue;
        }
        const [name, id] = line.split(',');

        if (!id) {
            console.error(`Invalid line: ${line}`);
            continue;
        }

        // Skip until we reach continueFromId, if set
        if (skipping) {
            if (Number(id.trim()) !== continueFromId) {
                continue;
            } else {
                skipping = false;
            }
        }

        try {
            // Call first API with "subscription" header
            const testApiUrl = `https://www.example.com/testing/${id.trim()}`;
            const testApiResp = await fetch(testApiUrl, {
                method: 'POST',
                headers: {
                    'subscription': 'your-subscription-value'
                }
            });
            if (!testApiResp.ok) throw new Error(`Failed to fetch ${testApiUrl}: ${testApiResp.statusText}`);
            const testApiData: TestApiData = await testApiResp.json();

            const username = testApiData?.data?.username;
            if (!username) throw new Error(`Username not found in response for id ${id}`);

            // Call suspend API with "subscription" header
            const suspendApiUrl = `https://www.example.com/suspend/${username}`;
            const suspendApiResp = await fetch(suspendApiUrl, {
                method: 'POST',
                headers: {
                    'subscription': 'your-subscription-value'
                }
            });
            if (!suspendApiResp.ok) throw new Error(`Failed to suspend ${username}: ${suspendApiResp.statusText}`);
            const suspendApiData = await suspendApiResp.json();

            // Log to file
            const logEntry: LogEntry = {
                name: name.trim(),
                id: id.trim(),
                username,
                testApiResponse: testApiData,
                suspendApiResponse: suspendApiData
            };
            fs.appendFileSync(logFile, JSON.stringify(logEntry) + '\n');

            // Log to console
            console.log(`Processed id=${id.trim()}, username=${username}: Success`);
        } catch (err: any) {
            // Log error to file
            fs.appendFileSync(logFile, `Error processing id=${id}: ${err.message}\n`);
            console.error(`Processed id=${id.trim()}: Failed - ${err.message}`);
        }
    }
}

processCsv().catch(err => {
    console.error('Fatal error:', err);
});