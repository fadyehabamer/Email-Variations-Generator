function generateEmailVariations(email) {
    let [localPart, domain] = email.split("@");

    let variations = new Set();

    function addDots(prefix, remaining) {
        if (!remaining) {
            variations.add(prefix + "@" + domain);
            return;
        }

        // Without a dot
        addDots(prefix + remaining[0], remaining.slice(1));

        // With a dot
        if (remaining.length > 1) {
            addDots(prefix + remaining[0] + ".", remaining.slice(1));
        }
    }

    addDots("", localPart);

    return Array.from(variations);
}

let previouslyGeneratedEmail = '';

function generateAndDisplayVariations() {
    let emailInput = document.getElementById('emailInput');
    let email = emailInput.value;
    let errorMsg = document.getElementById('errorMsg');
    let exportBtn = document.getElementById('exportBtn');

    
    // Email validation
    if (!validateEmail(email)) {
        alert("Please enter a valid email address!");
        return;
    }

    // Check if the email is the same as the previously generated one
    if (email === previouslyGeneratedEmail) {
        alert("This email has already been generated!");
        return;
    }

    let resultsTable = document.getElementById('results');
    let validEmails = generateEmailVariations(email);

    // Clear previous results
    resultsTable.innerHTML = '';

    // Create and append rows and cells
    for (let i = 0; i < validEmails.length; i += 3) {
        let row = resultsTable.insertRow();

        for (let j = 0; j < 3; j++) {
            let cell = row.insertCell(j);
            if (validEmails[i + j]) {
                cell.textContent = validEmails[i + j];
            }
        }
    }

    // Enable the export button
    exportBtn.disabled = false;

    // Set this email as the previously generated one
    previouslyGeneratedEmail = email;

    // Disable the email input to prevent changing the email
    emailInput.disabled = true;
}

function validateEmail(email) {
    var re = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}$/;
    return re.test(email);
}

function exportToExcel() {
    let email = document.getElementById('emailInput').value;
    let validEmails = generateEmailVariations(email);

    let wb = XLSX.utils.book_new();
    let ws_name = "EmailVariations";
    let data = [["Main Email", "New Email"]];

    validEmails.forEach(variation => {
        data.push([email, variation]);
    });

    let ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, ws_name);
    XLSX.writeFile(wb, "email_variations.xlsx");
}
