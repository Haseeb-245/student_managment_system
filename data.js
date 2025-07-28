document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('studentForm');
    const tableBody = document.getElementById('studentTableBody');
    const fileSelectBtn = document.getElementById('fileSelectBtn');
    const saveBtn = document.getElementById('saveBtn');
    const statusDisplay = document.getElementById('statusDisplay');
    const submitBtn = document.getElementById('submitBtn');

    let students = [];
    let editIndex = null;
    let fileHandle = null;
    let isFileSelected = false;

    const restoreFileHandle = async () => {
        const handleData = localStorage.getItem('excelFileHandle');
        if (!handleData) return false;

        try {
            const { name, id } = JSON.parse(handleData);
            const handles = await window.showOpenFilePicker({
                types: [{
                    description: 'Excel Files',
                    accept: {
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
                    }
                }],
                multiple: false
            });

            for (const handle of handles) {
                const file = await handle.getFile();
                if (file.name === name && handle.id === id) {
                    fileHandle = handle;
                    if (await verifyFilePermissions()) {
                        await loadExcelFile();
                        isFileSelected = true;
                        saveBtn.disabled = false;
                        fileSelectBtn.style.display = 'none';
                        showStatus(`Working with: ${file.name}`, 'success');
                        return true;
                    }
                }
            }
        } catch (err) {
            console.error('Could not restore file handle:', err);
            showStatus('Could not restore previous file', 'error');
        }
        return false;
    };

    restoreFileHandle().then(restored => {
        if (!restored) {
            showStatus('Please select your Excel file', 'info');
        }
    });

    fileSelectBtn.addEventListener('click', async () => {
        try {
            fileSelectBtn.disabled = true;
            fileSelectBtn.innerHTML = 'Loading... <span class="spinner"></span>';

            [fileHandle] = await window.showOpenFilePicker({
                types: [{
                    description: 'Excel Files',
                    accept: {
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
                    }
                }],
                multiple: false
            });

            const file = await fileHandle.getFile();
            localStorage.setItem('excelFileHandle', JSON.stringify({
                name: file.name,
                id: fileHandle.id,
                lastModified: file.lastModified
            }));

            if (await verifyFilePermissions()) {
                await loadExcelFile();
                isFileSelected = true;
                saveBtn.disabled = false;
                fileSelectBtn.style.display = 'none';
                showStatus(`Working with: ${file.name}`, 'success');
            }
        } catch (err) {
            console.error('File selection error:', err);
            showStatus('File selection cancelled', 'error');
        } finally {
            fileSelectBtn.disabled = false;
            fileSelectBtn.textContent = 'Select Excel File';
        }
    });

    async function verifyFilePermissions() {
        const options = { mode: 'readwrite' };
        if (await fileHandle.queryPermission(options) === 'granted') return true;
        return await fileHandle.requestPermission(options) === 'granted';
    }

    async function loadExcelFile() {
        try {
            showStatus('Loading file...', 'info');
            const file = await fileHandle.getFile();
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            students = XLSX.utils.sheet_to_json(firstSheet);
            updateTable();
            showStatus(`Loaded ${students.length} students from ${file.name}`, 'success');
        } catch (err) {
            console.error('Load error:', err);
            showStatus('Error loading file', 'error');
        }
    }

    form.addEventListener('submit', function (e) {
        e.preventDefault();

        const student = {
            firstName: document.getElementById('firstName').value.trim(),
            lastName: document.getElementById('lastName').value.trim(),
            phone: document.getElementById('phone').value.trim(),
            email: document.getElementById('email').value.trim()
        };

        if (!validateStudent(student)) return;

        if (editIndex !== null) {
            students[editIndex] = student;
            editIndex = null;
            submitBtn.textContent = 'Add Student';
            showStatus('Student updated successfully', 'success');
        } else {
            students.push(student);
            showStatus('Student added successfully', 'success');
        }

        updateTable();
        form.reset();
    });

    function validateStudent(student) {
        if (!student.firstName || !student.lastName || !student.phone || !student.email) {
            showStatus('Please fill all fields', 'error');
            return false;
        }

        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(student.email)) {
            showStatus('Please enter a valid email address', 'error');
            return false;
        }

        if (!/^[\d\s\-()+]{8,}$/.test(student.phone)) {
            showStatus('Please enter a valid phone number', 'error');
            return false;
        }

        return true;
    }

    saveBtn.addEventListener('click', async (event) => {
        // ✅ prevent any default behavior (for safety)
const wb = XLSX.utils.book_new();
        event.preventDefault() 
        try {
            if (!(await verifyFilePermissions())) {
                showStatus('<i class="fas fa-exclamation-circle"></i> Need write permissions to save', 'error');
                return;
            }

            saveBtn.disabled = true;
            saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';

            
            const ws = XLSX.utils.json_to_sheet(students);
            XLSX.utils.book_append_sheet(wb, ws, "Students");

            const writable = await fileHandle.createWritable();
            await writable.write(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
            await writable.close();

            const file = await fileHandle.getFile();
            localStorage.setItem('excelFileHandle', JSON.stringify({
                name: file.name,
                id: fileHandle.id,
                lastModified: file.lastModified
            }));

            showStatus(`<i class="fas fa-check-circle"></i> Saved ${students.length} students to ${file.name} at ${new Date().toLocaleTimeString()}`, 'success');

          

          

        } catch (err) {
            console.error('Save error:', err);
            showStatus(`<i class="fas fa-exclamation-circle"></i> Error saving file: ${err.message}`, 'error');
            if (err.name === 'NotAllowedError') {
                fileSelectBtn.style.display = 'inline-block';
                saveBtn.disabled = true;
            }
        } 
    });

    function updateTable() {
        tableBody.innerHTML = '';

        if (students.length === 0) {
            const row = document.createElement('tr');
            row.innerHTML = `<td colspan="5" style="text-align: center; padding: 2rem;">No students found</td>`;
            tableBody.appendChild(row);
            return;
        }

        students.forEach((student, index) => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${student.firstName}</td>
                <td>${student.lastName}</td>
                <td>${formatPhone(student.phone)}</td>
                <td>${student.email}</td>
                <td>
                    <button class="action-btn btn-secondary" onclick="editStudent(${index})">Edit</button>
                    <button class="action-btn btn-danger" onclick="deleteStudent(${index})">Delete</button>
                </td>
            `;
            tableBody.appendChild(row);
        });
    }

    function formatPhone(phone) {
        const cleaned = ('' + phone).replace(/\D/g, '');
        const match = cleaned.match(/^(\d{3})(\d{3})(\d{4})$/);
        return match ? `(${match[1]}) ${match[2]}-${match[3]}` : phone;
    }

    function showStatus(message, type = 'info') {
        statusDisplay.textContent = message;
        statusDisplay.className = 'status-display';
        if (type) statusDisplay.classList.add(type);
    }

    window.editStudent = function (index) {
        const student = students[index];
        document.getElementById('firstName').value = student.firstName;
        document.getElementById('lastName').value = student.lastName;
        document.getElementById('phone').value = student.phone;
        document.getElementById('email').value = student.email;
        editIndex = index;
        submitBtn.textContent = 'Update Student';
        form.scrollIntoView({ behavior: 'smooth' });
        showStatus(`Editing student: ${student.firstName} ${student.lastName}`, 'info');
    };

    window.deleteStudent = function (index) {
        if (confirm(`Delete ${students[index].firstName} ${students[index].lastName} permanently?`)) {
            const deletedStudent = students.splice(index, 1)[0];
            updateTable();
            showStatus(`Deleted student: ${deletedStudent.firstName} ${deletedStudent.lastName} (remember to save)`, 'error');
        }
    };

  
});
