import puppeteer from 'puppeteer';
import fs from 'fs';
import ExcelJS from 'exceljs';
import readline from 'readline';

// Add this new function at the top of the file
function extractEmails(text) {
    try {
        // Regular expression to match email addresses
        const emailRegex = /[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}/g;
        
        // Extract all emails from the text
        const emails = text.match(emailRegex) || [];
        
        // Remove duplicates
        const uniqueEmails = [...new Set(emails)];
        
        console.log(`Found ${uniqueEmails.length} unique email addresses`);
        
        // Save to JSON file
        fs.writeFileSync('new_emails.json', JSON.stringify(uniqueEmails, null, 2));
        
        console.log('Emails saved to new_emails.json');
        return uniqueEmails;
    } catch (error) {
        console.error('Error extracting emails:', error);
        return [];
    }
}

async function extractEmailsToExcel(text) {
    try {
        const emailRegex = /[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}/g;
        const emails = text.match(emailRegex) || [];
        const uniqueEmails = [...new Set(emails)];
        
        fs.writeFileSync('new_emails.json', JSON.stringify(uniqueEmails, null, 2));
        
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Form Responses');

        // Define the validation lists to match exactly with the form
        const interactionLevels = ['1', '2', '3', '4', '5', '0'];
        const commitmentLevels = ['1', '2', '3', '4', '5', '0'];
        const behaviorChoices = ['Committed', 'Troublemaker', 'Absent'];

        // Add headers matching exactly with the form questions
        worksheet.columns = [
            { header: 'Student Email', key: 'email', width: 30 },
            { header: 'Teacher Code', key: 'teacherCode', width: 15 },
            { header: 'Student interaction level during the lecture, 1 is the lowest and 5 is the highest', key: 'interaction', width: 50 },
            { header: 'The level of student commitment during the session, 1 is the least committed and 5 is the most committed', key: 'commitment', width: 50 },
            { header: "What is this student's behavior:", key: 'behavior', width: 20 },
            { header: 'Feedback Or notes', key: 'feedback', width: 50 }
        ];

        // Add data validation for each column
        uniqueEmails.forEach((email, index) => {
            const rowNumber = index + 2;

            worksheet.addRow({
                email: email,
                teacherCode: 'AH-5352',
                interaction: '',
                commitment: '',
                behavior: '',
                feedback: ''
            });

            // Add data validation for interaction level
            worksheet.getCell(`C${rowNumber}`).dataValidation = {
                type: 'list',
                allowBlank: true,
                formulae: [`"${interactionLevels.join(',')}"`]
            };

            // Add data validation for commitment level
            worksheet.getCell(`D${rowNumber}`).dataValidation = {
                type: 'list',
                allowBlank: true,
                formulae: [`"${commitmentLevels.join(',')}"`]
            };

            // Add data validation for behavior
            worksheet.getCell(`E${rowNumber}`).dataValidation = {
                type: 'list',
                allowBlank: true,
                formulae: [`"${behaviorChoices.join(',')}"`]
            };
        });

        // Style the headers
        worksheet.getRow(1).font = { bold: true };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };

        // Add borders to all cells
        worksheet.eachRow((row) => {
            row.eachCell((cell) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });

        // Save the workbook
        await workbook.xlsx.writeFile('form_responses.xlsx');
        
        console.log(`Found ${uniqueEmails.length} unique email addresses`);
        console.log('Emails saved to new_emails.json and form_responses.xlsx');
        
        return uniqueEmails;
    } catch (error) {
        console.error('Error processing emails:', error);
        return [];
    }
}

// Example usage:

// const text = `your text containing emails here`;
// const emails = extractEmails(text);

const text = ` ADD UR STUDENT EMAILS HERE `;
await extractEmailsToExcel(text);

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

async function readUserInput(question) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    return new Promise(resolve => {
        rl.question(question, answer => {
            rl.close();
            resolve(answer);
        });
    });
}

async function submitFormsFromExcel() {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('form_responses.xlsx');
        const worksheet = workbook.getWorksheet('Form Responses');
        
        console.log('Excel file loaded successfully.');
        const answer = await readUserInput('Have you filled the Excel sheet? (y/N): ');
        
        if (answer.toLowerCase() !== 'y') {
            console.log('Please fill the Excel sheet and run the script again.');
            return;
        }

        const browser = await puppeteer.launch({
            headless: false,
            defaultViewport: null,
            args: ['--start-maximized']
        });

        try {
            const page = await browser.newPage();
            
            // Process each row (skipping header)
            for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
                const row = worksheet.getRow(rowNumber);
                const formData = {
                    email: row.getCell(1).text || '', // Use .text instead of .value
                    teacherCode: row.getCell(2).text || '',
                    interaction: row.getCell(3).text || '',
                    commitment: row.getCell(4).text || '',
                    behavior: row.getCell(5).text || '',
                    feedback: row.getCell(6).text || ''
                };

                if (!formData.email) continue; // Skip empty rows

                console.log(`Processing form for email: ${formData.email}`);

                try {
                    await page.goto('https://forms.office.com/Pages/ResponsePage.aspx?id=_hCKfMAMzkOyEhANlhpOOFsKvQyKu1BNjf_6-FtY2WJUMU5ZRVZHU0FVSklCSUVaREgxTUI3ME5CNC4u', {
                        waitUntil: 'networkidle0',
                        timeout: 30000
                    });

                    // Wait for form to load
                    await page.waitForSelector('input.-ar-67');
                    await page.waitForTimeout(2000);

                    // Fill email and teacher code
                    const inputs = await page.$$('input.-ar-67');
                    if (inputs.length >= 3) {
                        // Fill email
                        await inputs[0].evaluate(input => input.value = ''); // Clear first
                        await inputs[0].type(String(formData.email));
                        await page.waitForTimeout(500);

                        // Fill teacher code
                        await inputs[1].evaluate(input => input.value = ''); // Clear first
                        await inputs[1].type(String(formData.teacherCode));
                        await page.waitForTimeout(500);
                    }

                    // Get all radio buttons
                    const radioButtons = await page.$$('input[type="radio"]');

                    // Handle interaction level (Question 3)
                    if (formData.interaction) {
                        const interactionValue = parseInt(formData.interaction);
                        if (interactionValue >= 0 && interactionValue <= 5) {
                            let interactionIndex;
                            switch(interactionValue) {
                                case 1: interactionIndex = 0; break;
                                case 2: interactionIndex = 1; break;
                                case 3: interactionIndex = 2; break;
                                case 4: interactionIndex = 3; break;
                                case 5: interactionIndex = 4; break;
                                case 0: interactionIndex = 5; break;
                            }
                            
                            try {
                                await page.waitForTimeout(1000);
                                await page.evaluate((index) => {
                                    document.querySelectorAll('input[type="radio"]')[index].click();
                                }, interactionIndex);
                                console.log(`Selected interaction level: ${interactionValue}`);
                            } catch (error) {
                                console.error(`Failed to select interaction level: ${error.message}`);
                            }
                        }
                    }

                    // Handle commitment level (Question 4)
                    if (formData.commitment) {
                        const commitmentValue = parseInt(formData.commitment);
                        if (commitmentValue >= 0 && commitmentValue <= 5) {
                            let commitmentIndex;
                            switch(commitmentValue) {
                                case 1: commitmentIndex = 6; break;
                                case 2: commitmentIndex = 7; break;
                                case 3: commitmentIndex = 8; break;
                                case 4: commitmentIndex = 9; break;
                                case 5: commitmentIndex = 10; break;
                                case 0: commitmentIndex = 11; break;
                            }
                            
                            try {
                                await page.waitForTimeout(1000);
                                await page.evaluate((index) => {
                                    document.querySelectorAll('input[type="radio"]')[index].click();
                                }, commitmentIndex);
                                console.log(`Selected commitment level: ${commitmentValue}`);
                            } catch (error) {
                                console.error(`Failed to select commitment level: ${error.message}`);
                            }
                        }
                    }

                    // Handle behavior (Question 5)
                    if (formData.behavior) {
                        const behaviorMap = {
                            'Committed': 12,
                            'Troublemaker': 13,
                            'Absent': 14
                        };
                        const behaviorIndex = behaviorMap[formData.behavior];
                        if (behaviorIndex !== undefined) {
                            await radioButtons[behaviorIndex].evaluate(radio => radio.click());
                            await page.waitForTimeout(500);
                            console.log(`Selected behavior: ${formData.behavior}`);
                        }
                    }

                    // Fill feedback
                    if (formData.feedback && inputs.length >= 3) {
                        await inputs[2].evaluate(input => input.value = ''); // Clear first
                        await inputs[2].type(String(formData.feedback));
                        await page.waitForTimeout(500);
                    }

                    // Submit form with better handling
                    try {
                        // Wait for submit button with multiple possible selectors
                        const submitButton = await page.waitForSelector([
                            'button[type="submit"]',
                            '.office-form-bottom-button',
                            'button.css-221',
                            'button[data-automation-id="submitButton"]',
                            '.button-content'
                        ].join(','), {
                            visible: true,
                            timeout: 15000
                        });

                        // Scroll to the submit button
                        await submitButton.evaluate(button => {
                            button.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        });
                        await page.waitForTimeout(2000);

                        // Click using JavaScript click event
                        await page.evaluate(() => {
                            const submitButtons = document.querySelectorAll([
                                'button[type="submit"]',
                                '.office-form-bottom-button',
                                'button.css-221',
                                'button[data-automation-id="submitButton"]',
                                '.button-content'
                            ].join(','));
                            if (submitButtons.length > 0) {
                                submitButtons[0].click();
                            }
                        });

                        // Wait for submission to complete
                        await Promise.race([
                            page.waitForNavigation({ timeout: 5000 }),
                            page.waitForSelector('div[role="alert"]', { timeout: 5000 }),
                            page.waitForFunction(
                                () => document.body.innerText.includes('Thank you'),
                                { timeout: 30000 }
                            )
                        ]);

                        console.log(`Successfully submitted form for ${formData.email}`);
                        await page.waitForTimeout(5000);

                    } catch (submitError) {
                        console.error(`Error submitting form for ${formData.email}:`, submitError.message);
                        
                        // Take screenshot on error
                        await page.screenshot({
                            path: `error-${Date.now()}-${formData.email.replace('@', '_at_')}.png`,
                            fullPage: true
                        });

                        // Wait before continuing
                        await page.waitForTimeout(5000);
                    }

                    // Wait between submissions
                    await page.waitForTimeout(5000);
                } catch (error) {
                    console.error(`Error processing form for email: ${formData.email}:`, error.message);
                }
            } // End of row processing loop

        } finally {
            await browser.close();
        }

    } catch (error) {
        console.error('Error processing Excel file:', error);
    }
}

// Call the function
submitFormsFromExcel();
