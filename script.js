const { Document, Paragraph, TextRun, ImageRun, Packer, PageOrientation } = window.docx;

function createCardHTML() {
    return `
        <div class="card">
            <h5>Photo</h5>
            <input type="file" class="image-upload" accept="image/*" onchange="handleImageUpload(this)" required>
            <img src="" alt="Inspection Photo">
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

function addNewCard() {
    const container = document.getElementById('cardsContainer');
    container.insertAdjacentHTML('beforeend', createCardHTML());
}

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

async function getDocumentContent() {
    const { Paragraph, TextRun, ImageRun } = docx;
    const cards = document.querySelectorAll('.card');
    const content = [];

    try {
        for (const card of cards) {
            const sn = card.querySelector('.sn').value;
            const location = card.querySelector('.location').value;
            const comments = card.querySelector('.comments').value;
            const img = card.querySelector('img');

            if (img.src) {
                // Convert base64 image data
                const base64Data = img.src.split(',')[1];

                content.push(
                    new Paragraph({
                        children: [
                            new TextRun({ text: `S/N: ${sn}`, bold: true }),
                            new TextRun({ text: `\nLocation: ${location}` }),
                            new TextRun({ text: `\nComments: ${comments}` }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new ImageRun({
                                data: base64Data,
                                transformation: {
                                    width: 300,
                                    height: 200,
                                },
                                type: 'jpg',
                            }),
                        ],
                    }),
                    new Paragraph({ spacing: { after: 200 } })
                );
            }
        }
        return content;
    } catch (error) {
        console.error('Error generating document content:', error);
        throw error;
    }
}

async function exportToWord() {
    try {
        const cards = document.querySelectorAll('.card');
        if (cards.length === 0) {
            alert('No content to export');
            return;
        }

        document.getElementById('loadingOverlay').style.display = 'flex';

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        size: {
                            orientation: PageOrientation.PORTRAIT,
                        },
                        margin: {
                            top: 1000,
                            right: 1000,
                            bottom: 1000,
                            left: 1000,
                        },
                    },
                },
                children: []
            }]
        });

        // Process each card
        for (const card of cards) {
            const sn = card.querySelector('.sn').value;
            const location = card.querySelector('.location').value;
            const comments = card.querySelector('.comments').value;
            const img = card.querySelector('img');

            const paragraphs = [
                new Paragraph({
                    children: [
                        new TextRun({ text: `S/N: ${sn}`, bold: true, size: 24 }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: `Location: ${location}`, size: 24 }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: `Comments: ${comments}`, size: 24 }),
                    ],
                })
            ];

            // Add image if it exists
            if (img && img.src && img.src.startsWith('data:image')) {
                try {
                    const base64Data = img.src.split(',')[1];
                    paragraphs.push(
                        new Paragraph({
                            children: [
                                new ImageRun({
                                    data: base64Data,
                                    transformation: {
                                        width: 400,
                                        height: 300,
                                    },
                                    type: 'jpeg'
                                }),
                            ],
                        })
                    );
                } catch (imageError) {
                    console.error('Image processing error:', imageError);
                }
            }

            // Add spacing after each card
            paragraphs.push(new Paragraph({ spacing: { after: 400 } }));

            // Add all paragraphs to document
            doc.addSection({
                children: paragraphs
            });
        }

        // Generate and download
        const blob = await Packer.toBlob(doc);
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'inspection_report.docx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    } catch (error) {
        console.error('Export error:', error);
        alert(`Export failed: ${error.message}`);
    } finally {
        document.getElementById('loadingOverlay').style.display = 'none';
    }
}

// Initialize with one card
addNewCard();