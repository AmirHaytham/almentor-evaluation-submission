const puppeteer = require('puppeteer');
const fs = require('fs');

// Array of different compliments
const compliments = [
"",
""
];

async function wait(page, ms) {
    await page.evaluate(ms => new Promise(resolve => setTimeout(resolve, ms)), ms);
}

async function submitForms(emails) {
    let browser = await puppeteer.launch({
        headless: false,
        defaultViewport: null,
        args: ['--start-maximized']
    });

    try {
        let page = await browser.newPage();
        
        for (const email of emails) {
            let retryCount = 0;
            const maxRetries = 3;
            
            while (retryCount < maxRetries) {
                try {
                    console.log(`Processing form for email: ${email} (Attempt ${retryCount + 1})`);
                    
                    // Check if browser is still connected
                    if (!browser.isConnected()) {
                        console.log('Browser disconnected, launching new browser...');
                        browser = await puppeteer.launch({
                            headless: false,
                            defaultViewport: null,
                            args: ['--start-maximized']
                        });
                        page = await browser.newPage();
                        await wait(page, 3000);
                    }

                    // Close extra pages
                    const pages = await browser.pages();
                    for (const p of pages) {
                        if (p !== page) {
                            await p.close().catch(() => {});
                        }
                    }

                    // Navigate to the form with better error handling
                    let retries = 3;
                    while (retries > 0) {
                        try {
                            await page.goto('https://forms.office.com/Pages/ResponsePage.aspx?id=_hCKfMAMzkOyEhANlhpOOFsKvQyKu1BNjf_6-FtY2WJUMU5ZRVZHU0FVSklCSUVaREgxTUI3ME5CNC4u', {
                                waitUntil: 'networkidle0',
                                timeout: 30000
                            });
                            break;
                        } catch (error) {
                            retries--;
                            if (retries === 0) throw error;
                            console.log('Retrying page load...');
                            await wait(page, 5000);
                        }
                    }

                    // Wait for the form to be actually interactive
                    await page.waitForFunction(() => {
                        return document.querySelectorAll('input').length > 0;
                    }, { timeout: 15000 });

                    // Ensure the page is fully loaded
                    await wait(page, 5000);

                    // Get all input fields and verify they exist
                    const inputs = await page.$$('input.-ar-67');
                    if (!inputs || inputs.length < 3) {
                        throw new Error('Required input fields not found');
                    }

                    // Fill email with verification
                    await inputs[0].evaluate(el => el.value = ''); // Clear first
                    await inputs[0].type(email, { delay: 50 });
                    await wait(page, 1000);

                    // Fill teacher code with verification
                    await inputs[1].evaluate(el => el.value = ''); // Clear first
                    await inputs[1].type('AH-5352', { delay: 50 });
                    await wait(page, 1000);

                    // Handle radio buttons with explicit waiting
                    const radioButtons = await page.$$('input[type="radio"]');
                    if (radioButtons.length >= 14) {
                        await Promise.all([
                            radioButtons[4].evaluate(el => el.scrollIntoView()),
                            wait(page, 500)
                        ]);
                        await radioButtons[4].click();
                        await wait(page, 500);

                        await radioButtons[10].evaluate(el => el.scrollIntoView());
                        await radioButtons[10].click();
                        await wait(page, 500);

                        await radioButtons[12].evaluate(el => el.scrollIntoView());
                        await radioButtons[12].click();
                        await wait(page, 500);
                    } else {
                        throw new Error('Radio buttons not found');
                    }

                    // Fill feedback with verification
                    const randomCompliment = compliments[Math.floor(Math.random() * compliments.length)];
                    await inputs[2].evaluate(el => el.value = ''); // Clear first
                    await inputs[2].type(randomCompliment, { delay: 50 });
                    await wait(page, 1000);

                    // Submit form with better button detection
                    const submitButton = await page.waitForSelector([
                        'button[type="submit"]',
                        'button.office-form-bottom-button',
                        'button.css-221',
                        'button[data-automation-id="submitButton"]'
                    ].join(', '), { timeout: 10000, visible: true });

                    let isSubmitted = false;
                    
                    if (submitButton) {
                        await submitButton.evaluate(el => el.scrollIntoView());
                        await submitButton.click();
                        
                        // Enhanced submission verification
                        let attempts = 0;
                        const maxAttempts = 3;

                        while (!isSubmitted && attempts < maxAttempts) {
                            try {
                                // Wait for any of these success indicators
                                await Promise.race([
                                    page.waitForNavigation({ 
                                        timeout: 40000,
                                        waitUntil: 'networkidle0'
                                    }),
                                    page.waitForSelector('div[role="alert"]', { 
                                        timeout: 40000 
                                    }),
                                    page.waitForFunction(
                                        () => {
                                            return document.body.innerText.includes('Thank you') ||
                                                   document.body.innerText.includes('submitted') ||
                                                   window.location.href.includes('thankyou');
                                        },
                                        { timeout: 40000 }
                                    )
                                ]);

                                // Verify submission
                                const currentUrl = await page.url();
                                if (!currentUrl.includes('ResponsePage')) {
                                    isSubmitted = true;
                                    console.log(`Successfully submitted form for email: ${email}`);
                                } else {
                                    // Check if form is actually submitted despite being on same URL
                                    const pageContent = await page.content();
                                    if (pageContent.includes('Thank you') || 
                                        pageContent.includes('submitted')) {
                                        isSubmitted = true;
                                        console.log(`Successfully submitted form for email: ${email}`);
                                    }
                                }

                            } catch (submitError) {
                                attempts++;
                                console.log(`Submission attempt ${attempts} failed for ${email}`);
                                
                                if (attempts === maxAttempts) {
                                    throw new Error(`Failed to verify submission after ${maxAttempts} attempts`);
                                }
                                
                                // Check if we need to retry submission
                                const stillOnFormPage = await page.evaluate(() => {
                                    return document.querySelector('button[type="submit"]') !== null;
                                });

                                if (stillOnFormPage) {
                                    console.log('Still on form page, retrying submission...');
                                    await wait(page, 5000);
                                    await submitButton.click();
                                }
                            }
                        }

                        // Wait longer after successful submission
                        if (isSubmitted) {
                            await wait(page, 8000);
                            console.log(`Successfully submitted form for email: ${email}`);
                            // Close current page and create a new one for the next submission
                            await page.close();
                            page = await browser.newPage();
                            await wait(page, 3000); // Wait a bit before starting next submission
                        }
                    }

                    break; // Break the retry loop if successful
                    
                } catch (error) {
                    console.error(`Error (Attempt ${retryCount + 1}) processing email ${email}:`, error.message);
                    retryCount++;
                    
                    try {
                        // Try to close the page if it's still open
                        if (page && !page.isClosed()) {
                            await page.screenshot({ 
                                path: `error-${Date.now()}-${email.replace('@', '_at_')}.png` 
                            }).catch(() => {});
                            await page.close().catch(() => {});
                        }
                    } catch (e) {
                        console.log('Error while cleaning up:', e.message);
                    }

                    if (retryCount < maxRetries) {
                        console.log(`Retrying email ${email} (Attempt ${retryCount + 1} of ${maxRetries})`);
                        try {
                            // Restart browser for next attempt
                            await browser.close().catch(() => {});
                            browser = await puppeteer.launch({
                                headless: false,
                                defaultViewport: null,
                                args: ['--start-maximized']
                            });
                            page = await browser.newPage();
                            await wait(page, 5000);
                        } catch (browserError) {
                            console.error('Error restarting browser:', browserError.message);
                        }
                    } else {
                        console.error(`Failed to process email ${email} after ${maxRetries} attempts`);
                    }
                }
            }
        }

    } catch (error) {
        console.error('Fatal error:', error.message);
    } finally {
        try {
            await browser.close().catch(() => {});
        } catch (closeError) {
            console.error('Error closing browser:', closeError.message);
        }
    }
}

// Main function to run the script
async function main() {
    // Read emails from the text file
    const emailsText = fs.readFileSync('new_emails.json', 'utf8');
    const emails = JSON.parse(emailsText);
    
    // Find indices of start and end emails
    const startEmail = 'muhammed-eltohamy@hotmail.com';
    const endEmail = '472294892@sharkia4.moe.edu.eg';
    
    const startIndex = emails.indexOf(startEmail);
    const endIndex = emails.indexOf(endEmail);
    
    if (startIndex === -1 || endIndex === -1) {
        console.error('One or both of the specified emails not found in the list');
        console.log('Start email:', startEmail, 'found:', startIndex);
        console.log('End email:', endEmail, 'found:', endIndex);
        // Debug: Print first few emails from the list
        console.log('First few emails in the list:', emails.slice(0, 5));
        return;
    }
    
    // Extract emails in the specified range
    const emailsToProcess = emails.slice(
        Math.min(startIndex, endIndex),
        Math.max(startIndex, endIndex) + 1
    );
    
    console.log(`Starting form submissions for ${emailsToProcess.length} emails...`);
    console.log('Processing emails from:', startEmail);
    console.log('To:', endEmail);
    console.log('Emails to process:', emailsToProcess); // Debug line
    
    await submitForms(emailsToProcess);
}

main(); 