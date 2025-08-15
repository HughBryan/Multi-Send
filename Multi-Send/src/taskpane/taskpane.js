// Global state
let recipients = [];
let nextId = 1;

// Initialize the application
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Email Duplicator add-in loaded successfully");
        addRecipient(); // Start with one row
        updateCount();
        setupPasteListener();
        setupPasteArea();
    }
});

// Setup paste area functionality
function setupPasteArea() {
    const pasteArea = document.getElementById('pasteArea');
    
    pasteArea.addEventListener('blur', function() {
        const value = this.value.trim();
        if (value && value.includes('@')) {
            console.log('Processing paste area data:', value);
            processPasteData(value);
            this.value = ''; // Clear after processing
        }
    });
    
    pasteArea.addEventListener('paste', function(e) {
        // Small delay to ensure paste content is in the textarea
        setTimeout(() => {
            const value = this.value.trim();
            if (value && value.includes('@')) {
                console.log('Processing pasted data in paste area:', value);
                processPasteData(value);
                this.value = ''; // Clear after processing
            }
        }, 100);
    });
}

// Setup paste event listener
function setupPasteListener() {
    document.addEventListener('paste', function(e) {
        console.log('Paste event detected'); // Debug log
        handlePasteEvent(e);
    });
}

// Handle paste events
function handlePasteEvent(e) {
    let pasteData = '';
    
    // Try to get paste data from clipboard
    if (e.clipboardData && e.clipboardData.getData) {
        pasteData = e.clipboardData.getData('text/plain');
        console.log('Got clipboard data:', pasteData); // Debug log
    } else if (window.clipboardData && window.clipboardData.getData) {
        pasteData = window.clipboardData.getData('Text');
        console.log('Got IE clipboard data:', pasteData); // Debug log
    }

    // Check if we have email data
    if (pasteData && pasteData.includes('@')) {
        console.log('Processing email data'); // Debug log
        processPasteData(pasteData);
        e.preventDefault(); // Prevent default paste behavior
        return;
    }

    // Fallback: check target element after paste completes
    if (e.target && (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA')) {
        setTimeout(() => {
            const value = e.target.value;
            console.log('Checking input value:', value); // Debug log
            if (value && value.includes('@') && (value.includes('\n') || value.includes('\t'))) {
                processPasteData(value);
                e.target.value = ''; // Clear the input
            }
        }, 100);
    }
}

// Process pasted data
function processPasteData(pasteData) {
    const lines = pasteData.split('\n').filter(line => line.trim());
    const emailLines = lines.filter(line => line.includes('@'));
    
    if (emailLines.length === 0) return;

    showPasteHint();

    // Clear existing data if there's meaningful content
    const hasData = recipients.some(r => r.email.trim());
    if (hasData || recipients.length > 1) {
        clearAllRecipients();
    }

    // Parse and add emails
    let parsedCount = 0;
    lines.forEach(line => {
        const parsedData = parseEmailLine(line.trim());
        if (parsedData.email) {
            addRecipientWithData(parsedData.email, parsedData.name);
            parsedCount++;
        }
    });

    if (parsedCount === 0) {
        addRecipient(); // Add one empty row if nothing was parsed
    } else {
        showStatus(`Pasted ${parsedCount} email${parsedCount !== 1 ? 's' : ''} from clipboard`, 'success');
    }

    updateCount();
    hidePasteHint();
}

// Parse individual email line
function parseEmailLine(line) {
    let email = '';
    let name = '';

    // Try tab-separated first (most common from Excel)
    if (line.includes('\t')) {
        const parts = line.split('\t');
        email = parts[0].trim();
        name = parts[1] ? parts[1].trim() : '';
    }
    // Try comma-separated
    else if (line.includes(',')) {
        const parts = line.split(',');
        email = parts[0].trim();
        name = parts[1] ? parts[1].trim() : '';
    }
    // Try multiple spaces
    else if (line.includes('  ')) {
        const parts = line.split(/\s{2,}/);
        email = parts[0].trim();
        name = parts[1] ? parts[1].trim() : '';
    }
    // Single word - assume it's just an email
    else if (line.includes('@')) {
        email = line.trim();
    }

    // If no name provided, suggest one
    if (email && !name) {
        name = suggestNameFromEmail(email);
    }

    return { email, name };
}

// Suggest name from email address
function suggestNameFromEmail(email) {
    const localPart = email.split('@')[0];
    if (localPart) {
        let namePart = localPart;
        // Replace dots and underscores with spaces
        namePart = namePart.replace(/[._]/g, ' ');
        // Remove numbers
        namePart = namePart.replace(/\d+/g, '');
        // Capitalize first letter of each word
        namePart = namePart.split(' ')
            .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
            .join(' ');
        // Take just the first name for simplicity
        return namePart.split(' ')[0];
    }
    return '';
}

// Add recipient with data
function addRecipientWithData(email, name) {
    const recipient = {
        id: nextId++,
        email: email,
        name: name
    };
    recipients.push(recipient);
    renderRecipient(recipient);
}

// Add empty recipient
function addRecipient() {
    const recipient = {
        id: nextId++,
        email: '',
        name: ''
    };
    recipients.push(recipient);
    renderRecipient(recipient);
    updateCount();
}

// Render individual recipient row
function renderRecipient(recipient) {
    const emailList = document.getElementById('emailList');
    
    const row = document.createElement('div');
    row.className = 'email-input-row';
    row.id = `email-row-${recipient.id}`;
    
    row.innerHTML = `
        <input type="email" class="email-input" placeholder="recipient@company.com" 
               value="${recipient.email}" oninput="updateRecipientEmail(${recipient.id}, this.value)" />
        <input type="text" class="name-input" placeholder="Name" 
               value="${recipient.name}" oninput="updateRecipientName(${recipient.id}, this.value)" />
        <button class="suggest-btn" onclick="suggestName(${recipient.id})">Auto</button>
        <button class="remove-btn" onclick="removeRecipient(${recipient.id})" 
                ${recipients.length === 1 ? 'style="display:none"' : ''}>×</button>
    `;
    
    emailList.appendChild(row);
}

// Update recipient email
function updateRecipientEmail(id, value) {
    const recipient = recipients.find(r => r.id === id);
    if (recipient) {
        recipient.email = value;
        updateCount();
    }
}

// Update recipient name
function updateRecipientName(id, value) {
    const recipient = recipients.find(r => r.id === id);
    if (recipient) {
        recipient.name = value;
    }
}

// Remove recipient
function removeRecipient(id) {
    recipients = recipients.filter(r => r.id !== id);
    const row = document.getElementById(`email-row-${id}`);
    if (row) {
        row.remove();
    }
    updateCount();
    
    // Show remove buttons if more than one recipient
    updateRemoveButtons();
}

// Update visibility of remove buttons
function updateRemoveButtons() {
    const removeButtons = document.querySelectorAll('.remove-btn');
    removeButtons.forEach(btn => {
        btn.style.display = recipients.length > 1 ? 'inline-block' : 'none';
    });
}

// Clear all recipients
function clearAllRecipients() {
    recipients = [];
    nextId = 1;
    document.getElementById('emailList').innerHTML = '';
}

// Suggest name for specific recipient
function suggestName(id) {
    const recipient = recipients.find(r => r.id === id);
    if (recipient && recipient.email) {
        const suggestedName = suggestNameFromEmail(recipient.email);
        recipient.name = suggestedName;
        
        // Update the input field
        const nameInput = document.querySelector(`#email-row-${id} .name-input`);
        if (nameInput) {
            nameInput.value = suggestedName;
        }
    }
}

// Update recipient count display
function updateCount() {
    const filledRecipients = recipients.filter(r => r.email.trim());
    const count = filledRecipients.length;
    document.getElementById('countText').textContent = `${count} recipient${count !== 1 ? 's' : ''}`;
    
    // Update remove buttons visibility
    updateRemoveButtons();
}

// Show/hide paste hint
function showPasteHint() {
    document.getElementById('pasteHint').classList.add('show');
}

function hidePasteHint() {
    setTimeout(() => {
        document.getElementById('pasteHint').classList.remove('show');
    }, 1500);
}

// Detect placeholder in email
async function detectPlaceholder() {
    try {
        const result = await new Promise((resolve, reject) => {
            Office.context.mailbox.item.body.getAsync('html', (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(new Error('Failed to read email'));
                }
            });
        });

        const patterns = [
            /\{\{[\w\s]+\}\}/g,  // {{name}}, {{first name}}
            /\[[\w\s]+\]/g,      // [name], [first name]
            /\$[\w\s]+\$/g       // $name$, $first name$
        ];

        for (const pattern of patterns) {
            const matches = result.match(pattern);
            if (matches && matches.length > 0) {
                document.getElementById('placeholder').value = matches[0];
                showStatus(`Found placeholder: ${matches[0]}`, 'success');
                return;
            }
        }

        showStatus('No common placeholder patterns found. Using default {{name}}', 'info');
    } catch (error) {
        showStatus('Could not read email content', 'error');
    }
}

// Show status message
function showStatus(message, type = 'info') {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = `status ${type} show`;
    
    if (type === 'success') {
        setTimeout(() => {
            statusDiv.classList.remove('show');
        }, 4000);
    }
}

// Show/hide progress
function showProgress(show = true) {
    const progressDiv = document.getElementById('progress');
    progressDiv.classList.toggle('show', show);
    document.getElementById('generateBtn').disabled = show;
}

// Update progress
function updateProgress(current, total) {
    const percentage = total > 0 ? (current / total) * 100 : 0;
    document.getElementById('progressFill').style.width = `${percentage}%`;
    document.getElementById('progressText').textContent = `Creating email ${current} of ${total}...`;
}

// Get current email data
async function getCurrentEmailData() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync('html', (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                Office.context.mailbox.item.subject.getAsync((subjectResult) => {
                    if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
                        resolve({
                            body: result.value,
                            subject: subjectResult.value || '',
                            attachments: Office.context.mailbox.item.attachments || []
                        });
                    } else {
                        reject(new Error('Failed to get email subject'));
                    }
                });
            } else {
                reject(new Error('Failed to get email body'));
            }
        });
    });
}

// Create personalized email
async function createPersonalizedEmail(emailData, placeholder, recipient) {
    return new Promise((resolve) => {
        // Escape special regex characters in placeholder
        const escapedPlaceholder = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(escapedPlaceholder, 'g');
        
        const personalizedSubject = emailData.subject.replace(regex, recipient.name);
        const personalizedBody = emailData.body.replace(regex, recipient.name);

        const newEmail = {
            htmlBody: personalizedBody,
            subject: personalizedSubject,
            to: [recipient.email]
        };

        Office.context.mailbox.displayNewMessageForm(newEmail);
        setTimeout(() => resolve(), 100);
    });
}

// Generate personalized emails
async function generateEmails() {
    const placeholder = document.getElementById('placeholder').value.trim();
    const validRecipients = recipients.filter(r => r.email.trim() && r.name.trim());

    // Validation
    if (!placeholder) {
        showStatus('Please enter a placeholder text', 'error');
        return;
    }

    if (validRecipients.length === 0) {
        showStatus('Please add at least one recipient with email and name', 'error');
        return;
    }

    if (validRecipients.length > 15) {
        showStatus('Maximum 15 recipients allowed', 'error');
        return;
    }

    try {
        showProgress(true);
        showStatus('Reading current email...', 'info');

        // Get the current email content
        const currentEmail = await getCurrentEmailData();

        // Check if placeholder exists
        if (!currentEmail.body.includes(placeholder) && !currentEmail.subject.includes(placeholder)) {
            showStatus(`Placeholder "${placeholder}" not found in email`, 'error');
            showProgress(false);
            return;
        }

        // Generate personalized emails
        let successCount = 0;
        for (let i = 0; i < validRecipients.length; i++) {
            updateProgress(i + 1, validRecipients.length);

            try {
                await createPersonalizedEmail(currentEmail, placeholder, validRecipients[i]);
                successCount++;

                // Small delay to prevent overwhelming Outlook
                if (i < validRecipients.length - 1) {
                    await new Promise(resolve => setTimeout(resolve, 300));
                }
            } catch (error) {
                console.error(`Failed to create email for ${validRecipients[i].email}:`, error);
            }
        }

        showProgress(false);

        if (successCount === validRecipients.length) {
            showStatus(`Successfully created ${successCount} personalized emails!`, 'success');
        } else {
            showStatus(`Created ${successCount} of ${validRecipients.length} emails`, 'error');
        }

    } catch (error) {
        console.error('Error generating emails:', error);
        showStatus('Failed to generate emails. Please try again.', 'error');
        showProgress(false);
    }
}