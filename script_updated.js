// Biến toàn cục
let excelData = [];
let stream = null;

// DOM Elements
const excelFileInput = document.getElementById('excel-file');
const excelDataDiv = document.getElementById('excel-data');
const camera = document.getElementById('camera');
const canvas = document.getElementById('canvas');
const photo = document.getElementById('photo');
const resultDiv = document.getElementById('result');
const downloadLink = document.getElementById('download-link');

// Đọc file Excel
async function readExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

                const headers = jsonData[0].slice(0, 5); // chỉ lấy 5 cột đầu
                const rows = jsonData.slice(1);
                const result = rows.map(row => {
                    const obj = {};
                    headers.forEach((header, index) => {
                        obj[header] = row[index] || '';
                    });
                    return obj;
                });

                resolve(result);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('Lỗi khi đọc file'));
        reader.readAsArrayBuffer(file);
    });
}

// Hiển thị dữ liệu Excel
function displayExcelData(data) {
    if (data.length === 0) {
        excelDataDiv.innerHTML = '<p>Không có dữ liệu</p>';
        return;
    }

    const headers = Object.keys(data[0]);
    let html = '<table><thead><tr>';
    headers.forEach(h => { html += `<th>${h}</th>`; });
    html += '<th>Chụp ảnh</th></tr></thead><tbody>';

    data.forEach((row, index) => {
        html += '<tr>';
        headers.forEach(h => {
            html += `<td>${row[h]}</td>`;
        });
        html += `<td><button onclick="captureImage(${index})">📷</button></td></tr>`;
    });

    html += '</tbody></table>';
    excelDataDiv.innerHTML = html;
}

// Sự kiện khi chọn file Excel
excelFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const data = await readExcel(file);
        excelData = data;
        displayExcelData(data);
        initCamera();
    } catch (error) {
        alert('Lỗi khi đọc file Excel: ' + error.message);
    }
});

// Khởi tạo camera
async function initCamera() {
    try {
        stream = await navigator.mediaDevices.getUserMedia({
            video: {
                facingMode: 'environment',
                width: { ideal: 1280 },
                height: { ideal: 720 }
            }
        });
        camera.srcObject = stream;
    } catch (error) {
        alert('Không thể truy cập camera: ' + error.message);
    }
}

// Hàm chụp ảnh
function captureImage(index) {
    const row = excelData[index];
    const fileName = (row["Mã kks thiết bị"] || "anh") + ".jpg";

    // Chụp từ video
    canvas.width = camera.videoWidth;
    canvas.height = camera.videoHeight;
    const context = canvas.getContext('2d');
    context.drawImage(camera, 0, 0, canvas.width, canvas.height);

    // Hiển thị ảnh
    const imageUrl = canvas.toDataURL("image/jpeg", 0.9);
    photo.src = imageUrl;
    resultDiv.style.display = "block";

    // Tạo link tải
    downloadLink.href = imageUrl;
    downloadLink.download = fileName;
    downloadLink.textContent = `Tải về ${fileName}`;
    downloadLink.style.display = "block";
}
