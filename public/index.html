<!-- public/index.html -->

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <!-- If your content includes Hebrew or other non-ASCII characters, ensure UTF-8 encoding -->
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document & Image Processor</title>
    <!-- Bootstrap CSS -->
    <link
      rel="stylesheet"
      href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css"
      integrity="sha384-ME2dN43ZsCZlPSx37xZtB0IK1ScUyVezDlsMkXU0jJGnXkvv7VbpfbUH3Y6Xm0Rd"
      crossorigin="anonymous"
    />
    <link rel="stylesheet" href="styles.css">
    <style>
        /* Additional custom styles */
        .dropzone {
            border: 2px dashed #007bff;
            border-radius: 5px;
            padding: 40px;
            cursor: pointer;
            transition: background-color 0.3s, border-color 0.3s;
            margin-bottom: 20px;
            text-align: center;
            color: #6c757d;
        }

        .dropzone.dragover {
            background-color: #e9ecef;
            border-color: #007bff;
        }

        .dropzone label {
            font-size: 18px;
            font-weight: 500;
        }

        #upload-message {
            margin-top: 10px;
            color: green;
            font-weight: bold;
        }

        .btn-primary {
            width: 100%;
        }
    </style>
</head>
<body>
    <div class="container mt-5">
        <h1 class="text-center mb-4">Document & Image Processor</h1>
        <form action="/upload" method="POST" enctype="multipart/form-data">
            <div class="dropzone" id="dropzone">
                <input type="file" id="file-input" name="files" multiple accept=".pdf, .jpg, .jpeg, .png, .tiff, .tif">
                <label for="file-input">Tap to Select Files</label>
            </div>
            <div id="upload-message" class="text-center"></div>
            <button type="submit" class="btn btn-primary mt-3">Start Processing</button>
        </form>
    </div>

    <!-- Bootstrap JS and dependencies -->
    <script
      src="https://code.jquery.com/jquery-3.5.1.slim.min.js"
      integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj"
      crossorigin="anonymous"
    ></script>
    <script
      src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.bundle.min.js"
      integrity="sha384-LtrjvnR4/JWT+Fn6Jn3AK8C3F1O5Zy5eN1Vf1A5N09l2ZDbJhP6d2aW1UygOqNQu"
      crossorigin="anonymous"
    ></script>

    <script>
        const fileInput = document.getElementById('file-input');
        const uploadMessage = document.getElementById('upload-message');
        const dropzone = document.getElementById('dropzone');

        fileInput.addEventListener('change', () => {
            const files = Array.from(fileInput.files);
            if (files.length > 0) {
                // Filter accepted file types for user feedback
                const acceptedTypes = ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.tif'];
                const validFiles = files.filter(file => {
                    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
                    return acceptedTypes.includes(ext);
                });
                uploadMessage.textContent = `Selected ${validFiles.length} supported file(s). Ready to process.`;
            } else {
                uploadMessage.textContent = '';
            }
        });

        // Only add drag-and-drop functionality if not on a mobile device
        if (!/Mobi|Android/i.test(navigator.userAgent)) {
            // Enhance dropzone interactivity
            dropzone.addEventListener('dragover', (e) => {
                e.preventDefault();
                dropzone.classList.add('dragover');
            });

            dropzone.addEventListener('dragleave', () => {
                dropzone.classList.remove('dragover');
            });

            dropzone.addEventListener('drop', (e) => {
                e.preventDefault();
                dropzone.classList.remove('dragover');
                const files = Array.from(e.dataTransfer.files);
                if (files.length > 0) {
                    const dataTransfer = new DataTransfer();
                    files.forEach(file => dataTransfer.items.add(file));
                    fileInput.files = dataTransfer.files;

                    const acceptedTypes = ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.tif'];
                    const validFiles = files.filter(file => {
                        const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
                        return acceptedTypes.includes(ext);
                    });
                    uploadMessage.textContent = `Selected ${validFiles.length} supported file(s). Ready to process.`;
                }
            });
        }
    </script>
</body>
</html>
