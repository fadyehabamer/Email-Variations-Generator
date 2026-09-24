// The number of variations doubles with every character of the username
// (2^(n-1)), so a long username would freeze the page. Cap the output; 8192
// covers every variation of usernames up to 14 characters.
const MAX_VARIATIONS = 8192;

// Total number of dot variations for an email (may exceed MAX_VARIATIONS).
function countEmailVariations(email) {
    let localPart = email.split("@")[0].split("+")[0].replace(/\./g, "");
    return 2 ** Math.max(localPart.length - 1, 0);
}

function generateEmailVariations(email, limit = MAX_VARIATIONS) {
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
        if (variations.size >= limit) return;

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

// Email and variations currently shown in the table (used by the export).
let currentEmail = '';
let currentVariations = [];

function generateAndDisplayVariations() {
    let emailInput = document.getElementById('emailInput');
    let email = emailInput.value.trim();
    let errorMsg = document.getElementById('errorMsg');
    let statusMsg = document.getElementById('statusMsg');
    let exportBtn = document.getElementById('exportBtn');

    
    // Email validation
    if (!validateEmail(email)) {
        errorMsg.textContent = "Please enter a valid email address.";
        emailInput.setAttribute('aria-invalid', 'true');
        emailInput.focus();
        return;
    }
    errorMsg.textContent = '';
    emailInput.removeAttribute('aria-invalid');

    let resultsTable = document.getElementById('results');
    let validEmails = generateEmailVariations(email);
    let total = countEmailVariations(email);
    statusMsg.textContent = total > validEmails.length
        ? `Showing the first ${validEmails.length.toLocaleString()} of ${total.toLocaleString()} possible variations.`
        : `${validEmails.length.toLocaleString()} variations generated.`;

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

    currentEmail = email;
    currentVariations = validEmails;
}

function validateEmail(email) {
    // Local part: dot-separated atoms (no leading, trailing or double dots),
    // "+" allowed for tags. Domain: dot-separated labels that don't start or
    // end with "-", followed by a TLD of 2+ letters (e.g. ".photography").
    var re = /^[a-zA-Z0-9_%+-]+(\.[a-zA-Z0-9_%+-]+)*@([a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}$/;
    return re.test(email);
}

function exportToExcel() {
    // Export exactly what is displayed, even if the input was edited since.
    let email = currentEmail;
    let validEmails = currentVariations;
    if (!validEmails.length) return;

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
