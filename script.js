document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('registrationForm');
    const submitBtn = document.getElementById('submitBtn');
    const btnText = submitBtn.querySelector('.btn-text');
    const loader = submitBtn.querySelector('.loader');
    const messageDiv = document.getElementById('message');
    const fileInput = document.getElementById('screenshot');
    const fileNameDisplay = document.getElementById('file-name');

    // Apps Script URL - User to replace this
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyYyjnVuAnc8MHJ-iOMNXkUdeKI29FZn1IEwDs0e--ksX3V2qva9lXZAH4MCk6UeEZN/exec';

    // File input change handler to show filename
    fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length > 0) {
            fileNameDisplay.textContent = e.target.files[0].name;
            fileNameDisplay.style.color = '#1f2937';
        } else {
            fileNameDisplay.textContent = 'Choose File...';
            fileNameDisplay.style.color = '#6b7280';
        }
    });

    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        // Reset message
        messageDiv.style.display = 'none';
        messageDiv.className = 'message';
        messageDiv.textContent = '';

        // --- VALIDATION START ---

        // 1. Name (HTML5 handles required, but let's be sure)
        const name = document.getElementById('name').value.trim();
        if (!name) {
            showMessage('Please enter your full name.', 'error');
            return;
        }

        // 2. Email Validation
        const email = document.getElementById('email').value.trim();
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(email)) {
            showMessage('Please enter a valid email address.', 'error');
            return;
        }

        // 3. Phone Validation (Exactly 10 digits)
        const phone = document.getElementById('phone').value.trim();
        const phoneRegex = /^\d{10}$/;
        if (!phoneRegex.test(phone)) {
            showMessage('Phone number must be exactly 10 digits.', 'error');
            return;
        }

        // 4. Transaction ID Validation (Exactly 12 digits)
        const transactionId = document.getElementById('transactionId').value.trim();
        const txnRegex = /^\d{12}$/;
        if (!txnRegex.test(transactionId)) {
            showMessage('Transaction ID must be exactly 12 digits.', 'error');
            return;
        }

        // 5. File Validation (JPG/PNG only)
        const file = fileInput.files[0];
        if (!file) {
            showMessage('Please upload the payment screenshot.', 'error');
            return;
        }

        const allowedTypes = ['image/jpeg', 'image/png', 'image/jpg'];
        if (!allowedTypes.includes(file.type)) {
            showMessage('Only JPG and PNG images are allowed.', 'error');
            return;
        }

        // --- VALIDATION END ---

        if (APPS_SCRIPT_URL.includes('YOUR_APPS_SCRIPT_WEB_APP_URL_HERE')) {
            console.warn("Apps Script URL is not set.");
            // Proceed for simulation check later
        }

        // Set loading state
        setLoading(true);

        try {
            // Convert image to Base64
            const base64Image = await convertToBase64(file);

            // Prepare payload
            const formData = {
                name: name,
                email: email,
                phone: phone,
                transactionId: transactionId, // Changed from notes
                screenshot: base64Image,
                mimeType: file.type,
                fileName: file.name
            };

            console.log('Submitting to:', APPS_SCRIPT_URL);

            const response = await fetch(APPS_SCRIPT_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8',
                },
                body: JSON.stringify(formData)
            });

            console.log('Response status:', response.status);
            console.log('Response ok:', response.ok);

            if (!response.ok) {
                setLoading(false);
                throw new Error(`Server error: ${response.status}`);
            }

            // Try to parse response
            let result;
            try {
                const responseText = await response.text();
                console.log('Response text:', responseText);
                result = JSON.parse(responseText);
                console.log('Parsed result:', result);
            } catch (parseError) {
                console.error('Parse error:', parseError);
                setLoading(false);
                throw new Error('Invalid response from server');
            }

            // Check for success
            if (result.status === 'success' || result.result === 'success') {
                setLoading(false);
                alert('✅ Registration Successful! We will contact you soon.');
                showMessage('✅ Registration Successful! We will contact you soon.', 'success');
                form.reset();
                fileNameDisplay.textContent = 'Click to Upload Screenshot';
            } else {
                setLoading(false);
                alert('Error: ' + (result.message || 'Registration failed. Please try again.'));
                showMessage(result.message || 'Registration failed. Please try again.', 'error');
            }

        } catch (error) {
            console.error('Submission Error:', error);
            setLoading(false);
            alert('Error: ' + error.message);
            showMessage('Failed to submit: ' + error.message, 'error');
        }
    });

    function convertToBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => {
                // Remove the Data-URI prefix (e.g. "data:image/png;base64,") if just the raw base64 is needed.
                // However, usually it's better to keep it or let the backend parse it.
                // The prompt asked "Convert the file to base64", usually implies the string.
                // We'll send the full Data URL as it's easier to handle universaly.
                resolve(reader.result);
            };
            reader.onerror = error => reject(error);
        });
    }

    function setLoading(isLoading) {
        if (isLoading) {
            submitBtn.disabled = true;
            btnText.hidden = true;
            loader.hidden = false;
        } else {
            submitBtn.disabled = false;
            btnText.hidden = false;
            loader.hidden = true;
        }
    }

    function showMessage(msg, type) {
        messageDiv.textContent = msg;
        messageDiv.classList.remove('success', 'error');
        messageDiv.classList.add(type);
        messageDiv.style.display = 'block';

        // Scroll message into view
        setTimeout(() => {
            messageDiv.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }, 100);
    }
});
