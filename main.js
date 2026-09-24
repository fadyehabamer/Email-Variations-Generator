function generateEmailVariations(email) {
    let [localPart, domain] = email.split("@");

    // Gmail ignores dots in the username, so start from the dot-free form.
    // Otherwise existing dots get combined with inserted ones and produce
    // invalid addresses such as "john..doe@gmail.com".
    // A "+tag" suffix is kept as-is; only the username before it is varied.
    let plusIndex = localPart.indexOf("+");
    let tag = plusIndex === -1 ? "" : localPart.slice(plusIndex);
    localPart = (plusIndex === -1 ? localPart : localPart.slice(0, plusIndex)).replace(/\./g, "");

    let variations = new Set();

    function addDots(prefix, remaining) {
        if (!remaining) {
            variations.add(prefix + tag + "@" + domain);
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
    // Local part: dot-separated atoms (no leading, trailing or double dots),
    // "+" allowed for tags. Domain: dot-separated labels that don't start or
    // end with "-", followed by a TLD of 2+ letters (e.g. ".photography").
    var re = /^[a-zA-Z0-9_%+-]+(\.[a-zA-Z0-9_%+-]+)*@([a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}$/;
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
