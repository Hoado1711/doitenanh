// Biến toàn cục
let excelData = [];
let currentRow = null;
let stream = null;

// DOM Elements
const excelFileInput = document.getElementById('excel-file');
const excelDataDiv = document.getElementById('excel-data');
const camera = document.getElementById('camera');
const canvas = document.getElementById('canvas');
const photo = document.getElementById('photo');
const captureBtn = document.getElementById('capture-btn');
const retakeBtn = document.getElementById('retake-btn');
const uploadBtn = document.getElementById('upload-btn');
const resultDiv = document.getElementById('result');
const downloadLink = document.getElementById('download-link');

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
                
                // Chuyển đổi thành mảng các object
                const headers = jsonData[0];
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
        
        reader.onerror = () => {
            reject(new Error('Lỗi khi đọc file'));
        };
        
        reader.readAsArrayBuffer(file);
    });
}

// Hiển thị dữ liệu Excel dưới dạng bảng
function displayExcelData(data) {
    if (data.length === 0) {
        excelDataDiv.innerHTML = '<p>Không có dữ liệu</p>';
        return;
    }

    const headers = Object.keys(data[0]);
    let html = `<table><thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>`;
    
    data.forEach((row, index) => {
        html += `<tr data-index="${index}">${headers.map(h => `<td>${row[h]}</td>`).join('')}</tr>`;
    });
    
    html += '</tbody></table>';
    excelDataDiv.innerHTML = html;

    // Thêm sự kiện click cho các hàng
    const rows = excelDataDiv.querySelectorAll('tr[data-index]');
    rows.forEach(row => {
        row.addEventListener('click', () => {
            // Xóa class selected từ tất cả các hàng
            rows.forEach(r => r.classList.remove('selected'));
            
            // Thêm class selected cho hàng được chọn
            row.classList.add('selected');
            
            // Lưu hàng hiện tại
            currentRow = excelData[row.dataset.index];
            
            // Kích hoạt nút chụp ảnh
            captureBtn.disabled = false;
        });
    });
}

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

// Sự kiện chụp ảnh
captureBtn.addEventListener('click', () => {
    // Đặt kích thước canvas bằng kích thước video
    canvas.width = camera.videoWidth;
    canvas.height = camera.videoHeight;
    
    // Vẽ frame video vào canvas
    const context = canvas.getContext('2d');
    context.drawImage(camera, 0, 0, canvas.width, canvas.height);
    
    // Hiển thị ảnh đã chụp
    photo.src = canvas.toDataURL('image/png');
    resultDiv.style.display = 'block';
    
    // Ẩn nút chụp, hiện nút chụp lại
    captureBtn.disabled = true;
    retakeBtn.disabled = false;
});

// Sự kiện chụp lại
retakeBtn.addEventListener('click', () => {
    resultDiv.style.display = 'none';
    captureBtn.disabled = false;
    retakeBtn.disabled = true;
});

// Sự kiện lưu ảnh
uploadBtn.addEventListener('click', () => {
    if (!currentRow) {
        alert('Vui lòng chọn một hàng từ bảng Excel trước');
        return;
    }
    
    // Lấy tên cột đầu tiên làm tên file
    const fileName = Object.values(currentRow)[0] + '.png';
    
    // Tạo URL cho ảnh
    const imageUrl = canvas.toDataURL('image/png');
    
    // Cập nhật link tải về
    downloadLink.href = imageUrl;
    downloadLink.download = fileName;
    downloadLink.textContent = `Tải về ${fileName}`;
    
    // Hiển thị link tải về
    downloadLink.style.display = 'block';
    
    alert(`Ảnh đã được đổi tên thành ${fileName}. Nhấn vào link để tải về.`);
});