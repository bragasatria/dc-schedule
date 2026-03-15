import './bootstrap';
import '../css/app.css';

// Main page functionality
document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('dropzone-file');
    const fileInfo = document.getElementById('file-info');
    const fileNameDisplay = document.getElementById('file-name');
    const removeFileBtn = document.getElementById('remove-file');
    const form = document.getElementById('uploadForm');
    const submitBtn = document.getElementById('submitBtn');
    const btnText = document.getElementById('btnText');
    const btnLoading = document.getElementById('btnLoading');

    if (fileInput) {
        fileInput.addEventListener('change', function() {
            if (this.files && this.files.length > 0) {
                fileNameDisplay.textContent = this.files[0].name;
                fileInfo.classList.remove('hidden');
                fileInfo.classList.add('flex');
            }
        });
    }

    if (removeFileBtn) {
        removeFileBtn.addEventListener('click', function() {
            fileInput.value = '';
            fileInfo.classList.add('hidden');
            fileInfo.classList.remove('flex');
        });
    }

    // Drag and Drop functionality
    const dropzone = document.querySelector('label[for="dropzone-file"]');
    if (dropzone) {
        dropzone.addEventListener('dragover', function(e) {
            e.preventDefault();
            this.classList.add('drag-over');
        });

        dropzone.addEventListener('dragleave', function(e) {
            e.preventDefault();
            this.classList.remove('drag-over');
        });

        dropzone.addEventListener('drop', function(e) {
            e.preventDefault();
            this.classList.remove('drag-over');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                const file = files[0];
                if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
                    fileInput.files = files;
                    fileInput.dispatchEvent(new Event('change'));
                } else {
                    alert('Hanya file Excel (.xlsx atau .xls) yang diperbolehkan!');
                }
            }
        });
    }

    if (form) {
        form.addEventListener('submit', function() {
            // Show loading state
            submitBtn.disabled = true;
            submitBtn.classList.add('opacity-75', 'cursor-not-allowed');
            btnText.classList.add('hidden');
            btnLoading.classList.remove('hidden');

            // Reset after 5 seconds (as download usually starts quickly)
            setTimeout(() => {
                submitBtn.disabled = false;
                submitBtn.classList.remove('opacity-75', 'cursor-not-allowed');
                btnText.classList.remove('hidden');
                btnLoading.classList.add('hidden');

                // Optional: Clear form
                // fileInput.value = '';
                // fileInfo.classList.add('hidden');
            }, 5000);
        });
    }
}); 