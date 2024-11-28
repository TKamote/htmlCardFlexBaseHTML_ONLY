// Card Template Function
function createCardHTML() {
    return `
        <div class="card">
            <h5 class="photo">Photo</h5>
            <input type="file" class="image-upload" accept="image/*" onchange="handleImageUpload(this)" required>
            <img src="" alt="Inspection Photo" style="display: none;">
            <div class="input-group">
                <input type="text" class="input-field sn" placeholder="S/N" required>
                <div class="error-message"></div>
            </div>
            <div class="input-group">
                <input type="text" class="input-field location" placeholder="Location" required>
                <div class="error-message"></div>
            </div>
            <div class="input-group">
                <textarea class="input-field comments" placeholder="Comments" required></textarea>
                <div class="error-message"></div>
            </div>
        </div>
    `;
}

// Add new card
function addNewCard() {
    const container = document.getElementById('cardsContainer');
    container.insertAdjacentHTML('beforeend', createCardHTML());
}

// Form Validation
function validateForm() {
    const cards = document.querySelectorAll('.card');
    let isValid = true;

    cards.forEach(card => {
        const fields = card.querySelectorAll('.input-field');
        const img = card.querySelector('img');

        if (img.src === '' || img.style.display === 'none') {
            isValid = false;
            showError(card.querySelector('.image-upload'), 'Image is required');
        }

        fields.forEach(field => {
            if (!field.value.trim()) {
                isValid = false;
                showError(field, 'This field is required');
            } else {
                clearError(field);
            }
        });
    });

    return isValid;
}

// Error Handling Functions
function showError(element, message) {
    element.classList.add('error');
    const errorDiv = element.parentElement.querySelector('.error-message');
    if (errorDiv) errorDiv.textContent = message;
}

function clearError(element) {
    element.classList.remove('error');
    const errorDiv = element.parentElement.querySelector('.error-message');
    if (errorDiv) errorDiv.textContent = '';
}

// Image Upload Handler
function handleImageUpload(input) {
    const file = input.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const img = input.parentElement.querySelector('img');
            img.src = e.target.result;
            img.style.display = 'block';
        }
        reader.readAsDataURL(file);
    }
}

// Document Content Generation
async function getDocumentContent() {
    const { Paragraph, TextRun, ImageRun } = docx;
    const cards = document.querySelectorAll('.card');
    const content = [];

    try {
        for (const card of cards) {
            // ... [Content generation code remains the same]
        }
        return content;
    } catch (error) {
        console.error('Error generating document content:', error);
        throw error;
    }
}

// Export Function
async function exportToWord() {
    const loadingOverlay = document.getElementById('loadingOverlay');

    if (!validateForm()) {
        alert('Please fill in all required fields');
        return;
    }

    try {
        loadingOverlay.style.display = 'flex';
        const { Document, Packer } = docx;
        
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        size: {
                            orientation: 'portrait'
                        }
                    },
                    column: {
                        count: 2,
                        space: 708,
                        separate: true,
                    }
                },
                children: await getDocumentContent()
            }]
        });

        const blob = await Packer.toBlob(doc);
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        document.body.appendChild(a);
        a.style = 'display: none';
        a.href = url;
        a.download = 'site_inspection.docx';
        a.click();
        window.URL.revokeObjectURL(url);
    } catch (error) {
        console.error('Error exporting document:', error);
        alert('An error occurred while exporting the document');
    } finally {
        loadingOverlay.style.display = 'none';
    }
}

// Initialize with one card
addNewCard();