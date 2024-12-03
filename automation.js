import fs from 'fs';
import path from 'path';
import puppeteer from 'puppeteer';
import ExcelJS from 'exceljs';
import readline from 'readline';
import { fileURLToPath } from 'url';

// Get the directory name of the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Custom wait function to replace page.waitForTimeout
async function wait(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function extractEmails(text) {
    try {
        // Regular expression to match email addresses
        const emailRegex = /[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}/g;
        
        // Extract all emails from the text
        const emails = text.match(emailRegex) || [];
        
        // Remove duplicates
        const uniqueEmails = [...new Set(emails)];
        
        console.log(`Found ${uniqueEmails.length} unique email addresses`);
        
        // Ensure output directory exists
        const outputDir = path.join(__dirname, 'output');
        if (!fs.existsSync(outputDir)){
            fs.mkdirSync(outputDir);
        }
        
        // Save to JSON file in output directory
        const newEmailsPath = path.join(outputDir, 'new_emails.json');
        fs.writeFileSync(newEmailsPath, JSON.stringify(uniqueEmails, null, 2));
        
        console.log(`Emails saved to ${newEmailsPath}`);
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
        
        // Ensure directory exists
        const outputDir = path.join(__dirname, 'output');
        if (!fs.existsSync(outputDir)){
            fs.mkdirSync(outputDir);
        }

        const newEmailsPath = path.join(outputDir, 'new_emails.json');
        const formResponsesPath = path.join(outputDir, 'form_responses.xlsx');
        
        // Remove existing files to ensure fresh start
        try {
            if (fs.existsSync(newEmailsPath)) {
                fs.unlinkSync(newEmailsPath);
            }
            if (fs.existsSync(formResponsesPath)) {
                fs.unlinkSync(formResponsesPath);
            }
        } catch (cleanupError) {
            console.warn('Warning: Could not remove existing files', cleanupError);
        }
        
        fs.writeFileSync(newEmailsPath, JSON.stringify(uniqueEmails, null, 2));
        
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Form Responses', {
            properties: {
                tabColor: { argb: 'FFC0C0C0' }
            }
        });

        // Define the validation lists to match exactly with the form
        const interactionLevels = ['0', '1', '2', '3', '4', '5'];
        const commitmentLevels = ['0', '1', '2', '3', '4', '5'];
        const behaviorChoices = ['Committed', 'Troublemaker', 'Absent'];
        
        // Generate weeks from Week 1 to Week 20
        const weekChoices = Array.from({length: 20}, (_, i) => `Week ${i + 1}`);
        
        // Add headers matching exactly with the form questions
        worksheet.columns = [
            { header: 'Student Email', key: 'email', width: 30 },
            { header: 'Teacher Code', key: 'teacherCode', width: 15 },
            { header: 'Week', key: 'week', width: 15 },
            { 
                header: 'Student interaction level during the lecture, 1 is the lowest and 5 is the highest', 
                key: 'interaction', 
                width: 50 
            },
            { 
                header: 'The level of student commitment during the session, 1 is the least committed and 5 is the most committed', 
                key: 'commitment', 
                width: 50 
            },
            { 
                header: "What is this student's behavior:", 
                key: 'behavior', 
                width: 20 
            },
            { 
                header: 'Feedback Or notes', 
                key: 'feedback', 
                width: 50 
            },
            { 
                header: 'Question 3', 
                key: 'question3', 
                width: 50 
            },
            { 
                header: 'Week Question', 
                key: 'weekQuestion', 
                width: 20 
            }
        ];

        // Add data validation for each column
        uniqueEmails.forEach((email, index) => {
            const rowNumber = index + 2;

            worksheet.addRow({
                email: email,
                teacherCode: 'ENTER YOUR ID',
                week: '',
                interaction: '',
                commitment: '',
                behavior: '',
                feedback: '',
                question3: '',
                weekQuestion: ''
            });

            // Add dropdown for week selection
            const weekCell = worksheet.getCell(`C${rowNumber}`);
            const weekValidation = worksheet.dataValidations.add(`C${rowNumber}`, {
                type: 'list',
                allowBlank: true,
                formulae: [weekChoices.join(',')]
            });

            // Add dropdown for interaction level
            const interactionCell = worksheet.getCell(`D${rowNumber}`);
            const interactionValidation = worksheet.dataValidations.add(`D${rowNumber}`, {
                type: 'list',
                allowBlank: true,
                formulae: [interactionLevels.join(',')]
            });

            // Add dropdown for commitment level
            const commitmentCell = worksheet.getCell(`E${rowNumber}`);
            const commitmentValidation = worksheet.dataValidations.add(`E${rowNumber}`, {
                type: 'list',
                allowBlank: true,
                formulae: [commitmentLevels.join(',')]
            });

            // Add dropdown for behavior
            const behaviorCell = worksheet.getCell(`F${rowNumber}`);
            const behaviorValidation = worksheet.dataValidations.add(`F${rowNumber}`, {
                type: 'list',
                allowBlank: true,
                formulae: [behaviorChoices.join(',')]
            });
        });

        // Style the headers
        const headerRow = worksheet.getRow(1);
        headerRow.font = { bold: true, color: { argb: 'FF000000' } };
        headerRow.fill = {
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

        // Save the workbook with comprehensive error handling
        try {
            // Ensure directory exists
            const outputDir = path.join(__dirname, 'output');
            if (!fs.existsSync(outputDir)) {
                fs.mkdirSync(outputDir, { recursive: true });
            }

            // Write the file with additional options
            await workbook.xlsx.writeFile(formResponsesPath, {
                filename: formResponsesPath,
                useStyles: true,
                useSharedStrings: true
            });

            console.log(`Found ${uniqueEmails.length} unique email addresses`);
            console.log(`Emails saved to ${newEmailsPath} and ${formResponsesPath}`);
            
            // Verify file was written correctly
            const stats = fs.statSync(formResponsesPath);
            if (stats.size === 0) {
                throw new Error('Generated Excel file is empty');
            }
        } catch (writeError) {
            console.error('Error writing Excel file:', writeError);
            
            // Additional diagnostic logging
            console.error('Write Error Details:', {
                message: writeError.message,
                stack: writeError.stack,
                path: formResponsesPath
            });

            throw writeError;
        }

        return uniqueEmails;
    } catch (error) {
        console.error('Error processing emails:', error);
        return [];
    }
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
                        await wait(3000);
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
                            await wait(5000);
                        }
                    }

                    // Ensure output directory exists for screenshots
                    const outputDir = path.join(__dirname, 'output', 'screenshots');
                    if (!fs.existsSync(outputDir)){
                        fs.mkdirSync(outputDir, { recursive: true });
                    }

                    // Take screenshot on error
                    const takeErrorScreenshot = async (errorType) => {
                        const screenshotPath = path.join(outputDir, `${email}_${errorType}_error.png`);
                        await page.screenshot({ path: screenshotPath });
                        console.log(`Error screenshot saved to ${screenshotPath}`);
                    };

                    // Wait for the form to be actually interactive
                    await page.waitForFunction(() => {
                        return document.querySelectorAll('input').length > 0;
                    }, { timeout: 15000 });

                    // Ensure the page is fully loaded
                    await wait(5000);

                    // Verify and update selectors for form elements
                    const selectors = {
                        textInput: 'input[data-automation-id="textInput"]',
                        dropdownButton: 'div[role="button"][aria-haspopup="listbox"]',
                        interactionRadio: 'input[role="radio"][name="r60105d5c443d4547bb455e5cbceba3dd"]',
                        commitmentRadio: 'input[role="radio"][name="r3424766a67a14130a6a83d62a9bab8c4"]',
                        behaviorRadio: 'input[role="radio"][name="rfaf57e4af66643c1b882fd6dc16554cb"]',
                        submitButton: 'button[data-automation-id="submitButton"]'
                    };

                    // Get all input fields and verify they exist
                    const textInputs = await page.$$(selectors.textInput);
                    if (!textInputs || textInputs.length < 3) {
                        throw new Error('Required input fields not found');
                    }

                    // Fill email with verification
                    await textInputs[0].evaluate(el => el.value = ''); // Clear first
                    await textInputs[0].type(email, { delay: 50 });
                    await wait(1000);

                    // Fill teacher code with verification
                    await textInputs[1].evaluate(el => el.value = ''); // Clear first
                    await textInputs[1].type('AH-5352', { delay: 50 });
                    await wait(1000);

                    // Handle radio buttons with explicit waiting
                    const interactionRadios = await page.$$(selectors.interactionRadio);
                    if (interactionRadios.length >= 6) {
                        await Promise.all([
                            interactionRadios[4].evaluate(el => el.scrollIntoView()),
                            wait(500)
                        ]);
                        await interactionRadios[4].click();
                        await wait(500);

                        const allRadios = await page.$$('input[role="radio"]');
                        if (allRadios.length >= 12) {
                            await allRadios[10].evaluate(el => el.scrollIntoView());
                            await allRadios[10].click();
                            await wait(500);

                            await allRadios[12].evaluate(el => el.scrollIntoView());
                            await allRadios[12].click();
                            await wait(500);
                        } else {
                            throw new Error('Radio buttons not found');
                        }
                    } else {
                        throw new Error('Radio buttons not found');
                    }

                    // Fill feedback with verification
                    const randomCompliment = compliments[Math.floor(Math.random() * compliments.length)];
                    await textInputs[2].evaluate(el => el.value = ''); // Clear first
                    await textInputs[2].type(randomCompliment, { delay: 50 });
                    await wait(1000);

                    // Submit form with better button detection
                    const buttonElement = await page.evaluate(() => {
                        const button = document.querySelector(selectors.submitButton);
                        if (button) {
                            const rect = button.getBoundingClientRect();
                            return {
                                x: rect.left + rect.width / 2,
                                y: rect.top + rect.height / 2,
                                found: true
                            };
                        }
                        return { found: false };
                    });

                    let isSubmitted = false;
                    
                    if (buttonElement.found) {
                        await page.mouse.click(buttonElement.x, buttonElement.y);
                        
                        // Wait for submission confirmation
                        await page.waitForFunction(
                            () => document.body.innerText.includes('Thank you') ||
                                   document.body.innerText.includes('submitted') ||
                                   window.location.href.includes('thankyou'),
                            { timeout: 40000 }
                        );

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

                    }

                    break; // Break the retry loop if successful
                    
                } catch (error) {
                    console.error(`Error (Attempt ${retryCount + 1}) processing email ${email}:`, error.message);
                    retryCount++;
                    
                    try {
                        // Try to close the page if it's still open
                        if (page && !page.isClosed()) {
                            await takeErrorScreenshot('page_load');
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
                            await wait(5000);
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

async function submitFormsFromExcel() {
    try {
        // Ensure the output directory exists
        const outputDir = path.join(__dirname, 'output');
        const formResponsesPath = path.join(outputDir, 'form_responses.xlsx');

        // Check if file exists and is not empty
        if (!fs.existsSync(formResponsesPath)) {
            throw new Error(`Excel file not found at ${formResponsesPath}`);
        }

        // Read the workbook
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(formResponsesPath);
        const worksheet = workbook.getWorksheet('Form Responses');

        // Collect form data from worksheet
        const formDataList = [];
        worksheet.eachRow((row, rowNumber) => {
            // Skip header row
            if (rowNumber === 1) return;

            // Extract data from each column
            const formData = {
                email: row.getCell(1).value,
                teacherCode: row.getCell(2).value,
                week: row.getCell(3).value,
                interaction: row.getCell(4).value,
                commitment: row.getCell(5).value,
                behavior: row.getCell(6).value,
                feedback: row.getCell(7).value,
                question3: row.getCell(8).value,
                weekQuestion: row.getCell(9).value
            };

            // Only add rows with an email
            if (formData.email && typeof formData.email === 'string') {
                formDataList.push(formData);
            }
        });

        // Validate form data
        if (formDataList.length === 0) {
            throw new Error('No valid form data found in the Excel file');
        }

        console.log(`Loaded ${formDataList.length} form entries`);

        // Launch browser
        const browser = await puppeteer.launch({
            headless: false,
            defaultViewport: null,
            args: ['--start-maximized']
        });

        try {
            const page = await browser.newPage();
            
            // Helper function for delay
            function delay(time) {
                return new Promise(resolve => setTimeout(resolve, time));
            }

            // Function to clear input fields
            async function clearInputFields(page) {
                const inputs = await page.$$('input[data-automation-id="textInput"]');
                for (const input of inputs) {
                    await input.click({ clickCount: 3 }); // Select all text
                    await input.press('Backspace'); // Clear the input
                }
                console.log('Cleared all input fields');
            }

            // Process each form entry
            for (const formData of formDataList) {
                let formPage = null;
                try {
                    console.log(`Starting submission for email: ${formData.email}`);

                    // Create a new page for each form to prevent frame detachment
                    formPage = await browser.newPage();
                    await formPage.setDefaultNavigationTimeout(60000);
                    await formPage.setDefaultTimeout(60000);
                    await formPage.setBypassCSP(true);

                    // Navigate to form page
                    await formPage.goto('https://forms.office.com/Pages/ResponsePage.aspx?id=_hCKfMAMzkOyEhANlhpOOFsKvQyKu1BNjf_6-FtY2WJUMU5ZRVZHU0FVSklCSUVaREgxTUI3ME5CNC4u', {
                        waitUntil: 'networkidle0',
                        timeout: 60000
                    });

                    console.log('Form page loaded');

                    // Wait for form to load
                    await formPage.waitForSelector('input[data-automation-id="textInput"]', { timeout: 20000 });
                    await delay(2000);

                    // Find all text inputs
                    const textInputs = await formPage.$$('input[data-automation-id="textInput"]');
                    
                    // Fill email
                    if (textInputs.length > 0 && formData.email) {
                        await textInputs[0].type(String(formData.email));
                        await delay(1000);
                    }

                    // Fill teacher code
                    if (textInputs.length > 1 && formData.teacherCode) {
                        await textInputs[1].type(String(formData.teacherCode));
                        await delay(1000);
                    }

                    // Set Week to 'Week 3' for all submissions
                    const weekValue = 'Week 3';

                    // Select Week
                    const weekDropdown = await formPage.$('div[aria-haspopup="listbox"][role="button"]');
                    if (weekDropdown) {
                        console.log('Opening week dropdown');
                        await weekDropdown.click();
                        await formPage.waitForSelector('div[role="listbox"]');
                        await delay(1000);

                        // Select 'Week 3'
                        const weekOption = await formPage.evaluate(() => {
                            const options = Array.from(document.querySelectorAll('div[role="option"] span.text-format-content'));
                            const option = options.find(option => option.textContent.trim() === 'Week 3');
                            if (option) {
                                console.log('Found Week 3 option');
                                option.parentElement.click();
                                return true;
                            }
                            console.warn('Week 3 option not found');
                            return false;
                        });

                        if (weekOption) {
                            console.log('Successfully selected Week 3');
                        } else {
                            console.warn('Failed to select Week 3');
                        }
                        await delay(1000);
                    }

                    // Select Interaction Level
                    if (formData.interaction) {
                        const interactionRadios = await formPage.$$('input[role="radio"][name="r60105d5c443d4547bb455e5cbceba3dd"]');
                        const interactionIndex = parseInt(formData.interaction) - 1;
                        if (interactionIndex >= 0 && interactionIndex < interactionRadios.length) {
                            await interactionRadios[interactionIndex].click();
                            await delay(1000);
                        }
                    }

                    // Select Commitment Level
                    if (formData.commitment) {
                        const commitmentRadios = await formPage.$$('input[role="radio"][name="r3424766a67a14130a6a83d62a9bab8c4"]');
                        const commitmentIndex = parseInt(formData.commitment) - 1;
                        if (commitmentIndex >= 0 && commitmentIndex < commitmentRadios.length) {
                            await commitmentRadios[commitmentIndex].click();
                            await delay(1000);
                        }
                    }

                    // Select Behavior
                    if (formData.behavior) {
                        const behaviorMap = {
                            'Committed': 0,
                            'Troublemaker': 1,
                            'Absent': 2
                        };
                        const behaviorIndex = behaviorMap[formData.behavior];
                        if (behaviorIndex !== undefined) {
                            const behaviorRadios = await formPage.$$('input[role="radio"][name="rfaf57e4af66643c1b882fd6dc16554cb"]');
                            if (behaviorIndex < behaviorRadios.length) {
                                await behaviorRadios[behaviorIndex].click();
                                await delay(1000);
                            }
                        }
                    }

                    // Fill Feedback
                    if (textInputs.length > 2 && formData.feedback) {
                        await textInputs[2].type(String(formData.feedback));
                        await delay(1000);
                    }

                    // Fill Question 3
                    if (textInputs.length > 3 && formData.question3) {
                        await textInputs[3].type(String(formData.question3));
                        await delay(1000);
                    }

                    // Handle "Week" question specifically
                    if (formData.weekQuestion) {
                        // Try to find the specific input or dropdown for the Week question
                        const weekQuestionInput = await formPage.$('input[placeholder="Week"]');
                        if (weekQuestionInput) {
                            await weekQuestionInput.type(String(formData.weekQuestion));
                            await delay(1000);
                        } else {
                            // Alternative selector if needed
                            const weekQuestionDropdown = await formPage.$('div[role="combobox"][aria-label="Week Question"]');
                            if (weekQuestionDropdown) {
                                await weekQuestionDropdown.click();
                                await delay(1000);

                                // Find and select the specific week option
                                const weekOption = await formPage.evaluate((weekToSelect) => {
                                    const options = Array.from(document.querySelectorAll('div[role="listbox"] div[role="option"] span[aria-label]'));
                                    const matchingOption = options.find(opt => 
                                        opt.getAttribute('aria-label') === weekToSelect || 
                                        opt.getAttribute('aria-label').includes(weekToSelect)
                                    );
                                    return matchingOption ? options.indexOf(matchingOption) : -1;
                                }, formData.weekQuestion);

                                if (weekOption !== -1) {
                                    const weekOptions = await formPage.$$('div[role="listbox"] div[role="option"]');
                                    await weekOptions[weekOption].click();
                                    await delay(1000);
                                } else {
                                    console.warn(`Could not find week for Week Question: ${formData.weekQuestion}`);
                                    
                                    // Take screenshot to help diagnose week question selection issues
                                    await formPage.screenshot({
                                        path: path.join(__dirname, 'output', 'screenshots', `week_question_error_${formData.email.replace('@', '_at_')}_${Date.now()}.png`),
                                        fullPage: true
                                    });
                                }
                            } else {
                                console.warn(`Could not find input for Week Question: ${formData.weekQuestion}`);
                            }
                        }
                    }

                    // Submit form
                    const submitButton = await formPage.$('button[data-automation-id="submitButton"]');
                    if (submitButton) {
                        await submitButton.click();
                        
                        // Wait for submission confirmation
                        await formPage.waitForFunction(
                            () => document.body.innerText.includes('Thank you'),
                            { timeout: 10000 }
                        );

                        // Clear input fields
                        await clearInputFields(formPage);
                    }

                    // Close the page after submission
                    await formPage.close();

                    // Wait between submissions
                    await delay(5000);

                } catch (submissionError) {
                    console.error(`Error submitting form for ${formData.email}:`, submissionError);

                    // Take screenshot on error
                    await formPage.screenshot({
                        path: path.join(__dirname, 'output', 'screenshots', `submission_error_${formData.email.replace('@', '_at_')}_${Date.now()}.png`),
                        fullPage: true
                    });
                } finally {
                    // Ensure the page is closed to avoid memory leaks
                    if (formPage && !formPage.isClosed()) {
                        await formPage.close();
                    }
                }
            }
        } finally {
            await browser.close();
        }
    } catch (error) {
        console.error('Error in submitFormsFromExcel:', error);
        throw error;
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

async function initializeEmails() {
    const text = `ADD UR STUDENT EMAILS`;
    await extractEmailsToExcel(text);
}

async function main() {
    try {
        // Initialize emails and create Excel file
        await initializeEmails();
        
        // Submit forms from the generated Excel file
        await submitFormsFromExcel();
    } catch (error) {
        console.error('Error in main process:', error);
    }
}

// Call the main function
main();
