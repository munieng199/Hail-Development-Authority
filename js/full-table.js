// متغيرات عامة
let tableData = null;
let filteredData = null;
let currentPage = 1;
let pageSize = 25;
let sortColumn = null;
let sortDirection = 'asc';
const currentDateSerial = 45931; // 2025-10-01

// دالة تحويل تسلسلي إلى تاريخ
function excelSerialToDate(serial) {
    if (isNaN(serial) || serial === '' || serial === 'مستمرة' || /L\d+|A\d+/.test(serial)) {
        return serial || '-'; // الاحتفاظ بالقيمة الأصلية
    }
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    if (isNaN(date_info.getTime())) return serial || '-';
    const year = date_info.getFullYear();
    const month = String(date_info.getMonth() + 1).padStart(2, '0');
    const day = String(date_info.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`; // عرض كما في الملف
}

// تنسيق التاريخ للعرض (بدون تحويل إلى هجري)
function formatDate(dateString) {
    return dateString || '-'; // الاحتفاظ بالقيمة الأصلية
}

// عند تحميل الصفحة
document.addEventListener('DOMContentLoaded', function () {
    initializeTable();
    setupEventListeners();
});

// تهيئة الجدول
function initializeTable() {
    const storedData = localStorage.getItem('excelData');
    if (storedData) {
        try {
            tableData = JSON.parse(storedData);
            tableData.tasks.forEach(task => {
                task['تاريخ  بدء المهمه'] = excelSerialToDate(task['تاريخ  بدء المهمه']);
                task['التاريخ المتوقع لانهاء المهمة'] = excelSerialToDate(task['التاريخ المتوقع لانهاء المهمة']);
                task['التاريخ الفعلي لانتهاء المهمة'] = excelSerialToDate(task['التاريخ الفعلي لانتهاء المهمة']);

                // تطبيع اسم الإدارة
                if (task['الإدارة']) {
                    task['الإدارة'] = task['الإدارة'].trim()
                        .replace(/\s+/g, ' ')
                        .replace(/أ/g, 'ا')
                        .replace(/إ/g, 'ا');
                }

                const expectedSerial = parseFloat(task['التاريخ المتوقع لانهاء المهمة']);
                let status = task['الحالة'] || '';
                const actualDate = task['التاريخ الفعلي لانتهاء المهمة'];

                if (status.includes('التسليم اليوم')) {
                    status = 'متأخر';
                }

                if (!status || status === '-') {
                    if (!isNaN(expectedSerial) && expectedSerial < currentDateSerial && (!actualDate || actualDate === '-')) {
                        status = 'متأخر';
                    } else if (actualDate && actualDate !== '-') {
                        status = 'مكتمل';
                    } else {
                        status = 'جاري العمل';
                    }
                }

                task['الحالة'] = status;

                if (!task['نسبة التقدم']) {
                    task['نسبة التقدم'] = status.includes('مكتمل') ? 1 : status.includes('جاري') ? 0.5 : 0;
                }
            });
            filteredData = [...tableData.tasks];
            initializeFilters();

            // تطبيق الفلتر المحفوظ من لوحة المعلومات
            const statusFilter = localStorage.getItem('taskStatusFilter');
            if (statusFilter) {
                document.getElementById('statusFilter').value = statusFilter;
                applyFilters();
                // إزالة الفلتر بعد تطبيقه لتجنب تطبيقه مرة أخرى عند تحديث الصفحة
                localStorage.removeItem('taskStatusFilter');
            } else {
                renderTable();
            }

            updateResultsInfo();
        } catch (error) {
            console.error('خطأ في تحليل البيانات:', error);
            showNoDataState();
        }
    } else {
        showNoDataState();
    }
}

// إعداد مستمعي الأحداث
function setupEventListeners() {
    document.getElementById('searchInput').addEventListener('input', debounce(applyFilters, 300));
    document.getElementById('departmentFilter').addEventListener('change', applyFilters);
    document.getElementById('statusFilter').addEventListener('change', applyFilters);
    document.getElementById('responsibleFilter').addEventListener('change', applyFilters);
    document.getElementById('startDateFrom').addEventListener('change', applyFilters);
    document.getElementById('startDateTo').addEventListener('change', applyFilters);

    const progressFilter = document.getElementById('progressFilter');
    progressFilter.addEventListener('input', function () {
        document.getElementById('progressValue').textContent = this.value + '%';
        applyFilters();
    });

    document.getElementById('clearFilters').addEventListener('click', clearFilters);
    document.getElementById('resetFilters').addEventListener('click', clearFilters);

    document.getElementById('pageSize').addEventListener('change', function () {
        pageSize = this.value === 'all' ? filteredData.length : parseInt(this.value);
        currentPage = 1;
        renderTable();
        renderPagination();
    });

    document.getElementById('exportExcel').addEventListener('click', exportToExcel);
    document.getElementById('exportPDF').addEventListener('click', exportToPDF);
    document.getElementById('printTable').addEventListener('click', printTable);

    document.querySelectorAll('.sortable').forEach(header => {
        header.addEventListener('click', function () {
            const column = this.dataset.column;
            handleSort(column);
        });
    });
}

// تهيئة الفلاتر
function initializeFilters() {
    const departmentFilter = document.getElementById('departmentFilter');
    const departmentsSet = new Set();

    // استخراج الإدارات مع التنظيف واستبعاد الإدارة العامة للتميز المؤسسي
    tableData.tasks.forEach(task => {
        if (task['الإدارة']) {
            let deptName = task['الإدارة'].trim()
                .replace(/\s+/g, ' ') // توحيد المسافات
                .replace(/أ/g, 'ا') // توحيد الألف
                .replace(/إ/g, 'ا'); // توحيد الألف

            // استبعاد الإدارة العامة للتميز المؤسسي
            if (deptName && deptName !== '-' && deptName !== '' &&
                !deptName.includes('الادارة العامة للتميز') &&
                !deptName.includes('الإدارة العامة للتميز')) {
                departmentsSet.add(deptName);
            }
        }
    });

    // تحويل Set إلى مصفوفة وترتيبها
    const departments = Array.from(departmentsSet).sort();

    // إضافة جميع الإدارات إلى الفلتر
    departments.forEach(dept => {
        const option = document.createElement('option');
        option.value = dept;
        option.textContent = dept;
        departmentFilter.appendChild(option);
    });

    const responsibleFilter = document.getElementById('responsibleFilter');
    const responsiblePersonsSet = new Set();

    // استخراج جميع الأسماء من المهام مع فصل الأسماء المتعددة
    tableData.tasks.forEach(task => {
        if (task['المسؤول عن المهمه']) {
            const respText = task['المسؤول عن المهمه'].trim();
            if (respText && respText !== '-' && respText !== '') {
                // تقسيم الأسماء المتعددة (مفصولة بـ + أو ، أو /)
                const names = respText.split(/[\+،,\/]/);

                names.forEach(name => {
                    // تنظيف الاسم من الألقاب والمسافات الزائدة
                    let cleanName = name.trim()
                        .replace(/^(أ\.|م\.|د\.|أ |م |د |أ\s|م\s|د\s)/g, '') // إزالة الألقاب
                        .replace(/\s+/g, ' ') // توحيد المسافات
                        .replace(/أ/g, 'ا') // توحيد الألف
                        .replace(/إ/g, 'ا') // توحيد الألف
                        .trim();

                    if (cleanName && cleanName !== '-' && cleanName !== '') {
                        responsiblePersonsSet.add(cleanName);
                    }
                });
            }
        }
    });

    // تحويل Set إلى مصفوفة وترتيبها
    const responsiblePersons = Array.from(responsiblePersonsSet).sort();

    // إضافة جميع المسؤولين إلى الفلتر
    responsiblePersons.forEach(person => {
        const option = document.createElement('option');
        option.value = person;
        option.textContent = person;
        responsibleFilter.appendChild(option);
    });
}

// تطبيق الفلاتر
function applyFilters() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const departmentFilter = document.getElementById('departmentFilter').value;
    const statusFilter = document.getElementById('statusFilter').value;
    const responsibleFilter = document.getElementById('responsibleFilter').value;
    const startDateFrom = document.getElementById('startDateFrom').value;
    const startDateTo = document.getElementById('startDateTo').value;
    const progressFilter = parseInt(document.getElementById('progressFilter').value) / 100;

    filteredData = tableData.tasks.filter(task => {
        // فلتر البحث
        if (searchTerm && !task['الموضوع/المهمة'].toLowerCase().includes(searchTerm)) return false;

        // فلتر الإدارة - تطبيع البيانات قبل المقارنة
        if (departmentFilter) {
            const taskDept = task['الإدارة'] ? task['الإدارة'].trim()
                .replace(/\s+/g, ' ')
                .replace(/أ/g, 'ا')
                .replace(/إ/g, 'ا') : '';
            if (taskDept !== departmentFilter) return false;
        }

        // فلتر الحالة - معاملة "التسليم اليوم" كمتأخر
        if (statusFilter) {
            const taskStatus = task['الحالة'] || '';
            if (statusFilter === 'مكتمل' && !taskStatus.includes('مكتمل')) return false;
            if (statusFilter === 'متأخر' && !taskStatus.includes('متأخر') && !taskStatus.includes('التسليم اليوم')) return false;
            if (statusFilter === 'جاري العمل' && !taskStatus.includes('جاري') && taskStatus !== 'مستمرة') return false;
        }

        // فلتر المسؤول - البحث في الأسماء المنظفة
        if (responsibleFilter) {
            const respText = task['المسؤول عن المهمه'] || '';
            const names = respText.split(/[\+،,\/]/).map(name =>
                name.trim().replace(/^(أ\.|م\.|د\.|أ |م |د |أ\s|م\s|د\s)/g, '').trim()
            );
            if (!names.some(name => name.includes(responsibleFilter) || responsibleFilter.includes(name))) return false;
        }

        // فلتر التاريخ
        const taskStart = new Date(task['تاريخ  بدء المهمه']);
        if (startDateFrom && taskStart < new Date(startDateFrom)) return false;
        if (startDateTo && taskStart > new Date(startDateTo)) return false;

        // فلتر نسبة التقدم
        const taskProgress = parseFloat(task['نسبة التقدم']) || 0;
        if (taskProgress < progressFilter) return false;

        return true;
    });

    currentPage = 1;

    if (filteredData.length === 0) {
        showNoResults();
    } else {
        hideNoResults();
        renderTable();
        renderPagination();
    }

    updateResultsInfo();
}

// مسح الفلاتر
function clearFilters() {
    document.getElementById('searchInput').value = '';
    document.getElementById('departmentFilter').value = '';
    document.getElementById('statusFilter').value = '';
    document.getElementById('responsibleFilter').value = '';
    document.getElementById('startDateFrom').value = '';
    document.getElementById('startDateTo').value = '';
    document.getElementById('progressFilter').value = '0';
    document.getElementById('progressValue').textContent = '0%';

    // إزالة الفلتر المحفوظ من لوحة المعلومات
    localStorage.removeItem('taskStatusFilter');

    applyFilters();
}

// معالجة الترتيب
function handleSort(column) {
    if (sortColumn === column) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        sortColumn = column;
        sortDirection = 'asc';
    }

    document.querySelectorAll('.sortable').forEach(header => {
        header.classList.remove('sort-asc', 'sort-desc');
    });

    const currentHeader = document.querySelector(`[data-column="${column}"]`);
    currentHeader.classList.add(sortDirection === 'asc' ? 'sort-asc' : 'sort-desc');

    filteredData.sort((a, b) => {
        let valueA = a[column] || '';
        let valueB = b[column] || '';

        if (column.includes('تاريخ')) {
            valueA = new Date(valueA);
            valueB = new Date(valueB);
        }

        if (column === 'نسبة التقدم') {
            valueA = parseFloat(valueA) || 0;
            valueB = parseFloat(valueB) || 0;
        }

        if (valueA < valueB) return sortDirection === 'asc' ? -1 : 1;
        if (valueA > valueB) return sortDirection === 'asc' ? 1 : -1;
        return 0;
    });

    renderTable();
}

// عرض الجدول - محسّن للعرض على الشاشات الصغيرة
function renderTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    const startIndex = (currentPage - 1) * pageSize;
    const endIndex = pageSize === filteredData.length ? filteredData.length : startIndex + pageSize;
    const pageData = filteredData.slice(startIndex, endIndex);

    pageData.forEach((task, index) => {
        const row = document.createElement('tr');
        const taskName = task['الموضوع/المهمة'].toLowerCase();
        
        // إضافة classes للتظليل
        if (taskName.includes('efqm') || taskName.includes('iso')) row.classList.add('highlight-yellow');
        else if (taskName.includes('خطة صرف')) row.classList.add('highlight-orange');
        else if (taskName.includes('27001') || taskName.includes('الأمن السيبراني')) row.classList.add('highlight-blue');

        // إضافة data-label attributes لجميع الخلايا للعرض على الجوالات
        // ✅ التصحيح: إضافة جميع الخلايا حتى التي كانت مخفية
        row.innerHTML = `
            <td data-label="الموضوع/المهمة">
                <div style="max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${task['الموضوع/المهمة'] || '-'}">
                    ${highlightSearchTerm(task['الموضوع/المهمة'] || '-')}
                </div>
            </td>
            <td data-label="الإدارة">${task['الإدارة'] || '-'}</td>
            <td data-label="المسؤول">${task['المسؤول عن المهمه'] || '-'}</td>
            <td data-label="تاريخ البدء">${formatDate(task['تاريخ  بدء المهمه'])}</td>
            <td data-label="التاريخ المتوقع">${formatDate(task['التاريخ المتوقع لانهاء المهمة'])}</td>
            <td data-label="الحالة">${createStatusBadge(task['الحالة'])}</td>
            <td data-label="نسبة التقدم" class="progress-bar-cell">${createProgressBar(task['نسبة التقدم'])}</td>
            <td data-label="الإجراءات">
                <div class="action-buttons">
                    <button class="btn-action btn-view" onclick="viewTaskDetails(${startIndex + index})" title="عرض التفاصيل">
                        <i class="fas fa-eye"></i>
                        <span class="d-none d-sm-inline">عرض التفاصيل</span>
                        <span class="d-sm-none">عرض</span>
                    </button>
                </div>
            </td>
        `;
        
        // إضافة تأخير للأنيميشن
        row.style.animationDelay = `${index * 0.05}s`;
        tbody.appendChild(row);
    });

    renderPagination();
}

// إنشاء شارة الحالة
function createStatusBadge(status) {
    if (!status) return '<span class="status-badge">غير محدد</span>';

    let badgeClass = 'status-badge';
    if (status.includes('مكتمل')) badgeClass += ' status-completed';
    else if (status.includes('متأخر') || status.includes('التسليم اليوم')) badgeClass += ' status-delayed';
    else badgeClass += ' status-in-progress';

    return `<span class="${badgeClass}">${status}</span>`;
}

// إنشاء شريط التقدم
function createProgressBar(progress) {
    const percentage = Math.round((parseFloat(progress) || 0) * 100);

    return `
        <div class="progress-bar-small">
            <div class="progress-fill-small" style="width: ${percentage}%"></div>
        </div>
        <div class="progress-text">${percentage}%</div>
    `;
}

// تمييز مصطلح البحث
function highlightSearchTerm(text) {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    if (!searchTerm) return text;

    const regex = new RegExp(`(${searchTerm})`, 'gi');
    return text.replace(regex, '<span class="highlight">$1</span>');
}

// عرض ترقيم الصفحات
function renderPagination() {
    const pagination = document.getElementById('pagination');
    pagination.innerHTML = '';

    if (pageSize >= filteredData.length) return;

    const totalPages = Math.ceil(filteredData.length / pageSize);

    const prevButton = document.createElement('li');
    prevButton.className = `page-item ${currentPage === 1 ? 'disabled' : ''}`;
    prevButton.innerHTML = `<a class="page-link" href="#" onclick="changePage(${currentPage - 1})">السابق</a>`;
    pagination.appendChild(prevButton);

    const startPage = Math.max(1, currentPage - 2);
    const endPage = Math.min(totalPages, currentPage + 2);

    for (let i = startPage; i <= endPage; i++) {
        const pageButton = document.createElement('li');
        pageButton.className = `page-item ${i === currentPage ? 'active' : ''}`;
        pageButton.innerHTML = `<a class="page-link" href="#" onclick="changePage(${i})">${i}</a>`;
        pagination.appendChild(pageButton);
    }

    const nextButton = document.createElement('li');
    nextButton.className = `page-item ${currentPage === totalPages ? 'disabled' : ''}`;
    nextButton.innerHTML = `<a class="page-link" href="#" onclick="changePage(${currentPage + 1})">التالي</a>`;
    pagination.appendChild(nextButton);
}

// تغيير الصفحة
function changePage(page) {
    const totalPages = Math.ceil(filteredData.length / pageSize);
    if (page < 1 || page > totalPages) return;
    currentPage = page;
    renderTable();
}

// عرض تفاصيل المهمة
function viewTaskDetails(index) {
    const task = filteredData[index];
    const modalBody = document.getElementById('taskModalBody');
    modalBody.innerHTML = `
        <div class="task-detail-item">
            <span class="task-detail-label">الموضوع/المهمة:</span>
            <div class="task-detail-value">${task['الموضوع/المهمة'] || '-'}</div>
        </div>
        <div class="task-detail-item">
            <span class="task-detail-label">الإدارة:</span>
            <div class="task-detail-value">${task['الإدارة'] || '-'}</div>
        </div>
        <div class="task-detail-item">
            <span class="task-detail-label">المسؤول عن المهمة:</span>
            <div class="task-detail-value">${task['المسؤول عن المهمه'] || '-'}</div>
        </div>
        <div class="task-detail-item">
            <span class="task-detail-label">تاريخ بدء المهمة:</span>
            <div class="task-detail-value">${formatDate(task['تاريخ  بدء المهمه'])}</div>
        </div>
        <div class="task-detail-item">
            <span class="task-detail-label">التاريخ المتوقع لانتهاء المهمة:</span>
            <div class="task-detail-value">${formatDate(task['التاريخ المتوقع لانهاء المهمة'])}</div>
        </div>
        <div class="task-detail-item">
            <span class="task-detail-label">التاريخ الفعلي لانتهاء المهمة:</span>
            <div class="task-detail-value">${formatDate(task['التاريخ الفعلي لانتهاء المهمة'])}</div>
        </div>
        <div class="task-detail-item">
            <span class="task-detail-label">الحالة:</span>
            <div class="task-detail-value">${createStatusBadge(task['الحالة'])}</div>
        </div>
        <div class="task-detail-item">
            <span class="task-detail-label">نسبة التقدم:</span>
            <div class="task-detail-value">${createProgressBar(task['نسبة التقدم'])}</div>
        </div>
        <div class="task-detail-item">
            <span class="task-detail-label">النسبة المستهدفة:</span>
            <div class="task-detail-value">${(parseFloat(task['النسبة المستهدفة']) * 100 || 0).toFixed(0)}%</div>
        </div>
        <div class="task-detail-item">
            <span class="task-detail-label">ملاحظات:</span>
            <div class="task-detail-value">${task['ملاحظات (ان وجدت)'] || 'لا توجد ملاحظات'}</div>
        </div>
    `;
    const modal = new bootstrap.Modal(document.getElementById('taskModal'));
    modal.show();
}

// تحديث معلومات النتائج
function updateResultsInfo() {
    document.getElementById('currentResults').textContent = filteredData.length;
    document.getElementById('totalResults').textContent = filteredData.length;
}

// عرض حالة عدم وجود نتائج
function showNoResults() {
    const noDataState = document.getElementById('noDataState');
    const tableCard = document.querySelector('.table-card');
    
    if (noDataState) noDataState.style.display = 'block';
    if (tableCard) tableCard.style.display = 'none';
}

// إخفاء حالة عدم وجود نتائج
function hideNoResults() {
    const noDataState = document.getElementById('noDataState');
    const tableCard = document.querySelector('.table-card');
    
    if (noDataState) noDataState.style.display = 'none';
    if (tableCard) tableCard.style.display = 'block';
}

// عرض حالة عدم وجود بيانات
function showNoDataState() {
    const loadingState = document.getElementById('loadingState');
    const filtersCard = document.querySelector('.filters-card');
    const tableCard = document.querySelector('.table-card');
    const noDataState = document.getElementById('noDataState');
    const noDataTitle = document.querySelector('.no-data-content h3');
    const noDataMessage = document.querySelector('.no-data-content p');
    
    if (loadingState) loadingState.style.display = 'none';
    if (filtersCard) filtersCard.style.display = 'none';
    if (tableCard) tableCard.style.display = 'none';
    if (noDataState) noDataState.style.display = 'block';
    if (noDataTitle) noDataTitle.textContent = 'لا توجد بيانات';
    if (noDataMessage) noDataMessage.textContent = 'يرجى العودة إلى الصفحة الرئيسية وتحميل ملف البيانات أولاً';
}

// تصدير إلى Excel
function exportToExcel() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(filteredData);
    XLSX.utils.book_append_sheet(wb, ws, 'المهام');
    XLSX.writeFile(wb, 'مهام_الهيئة.xlsx');
}

// تصدير إلى PDF
function exportToPDF() {
    window.print();
}

// طباعة الجدول
function printTable() {
    window.print();
}

// العودة للرئيسية
function goBack() {
    window.location.href = 'index.html';
}

// الانتقال إلى لوحة المعلومات
function goToDashboard() {
    window.location.href = 'dashboard.html';
}

// دالة التأخير للبحث
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// ✅ دالة إضافية لتحسين عرض البطاقات على الشاشات الصغيرة
function enhanceMobileCards() {
    // هذه الدالة ستضمن أن البطاقات تعمل بشكل صحيح على الجوال
    const isMobile = window.innerWidth <= 992;
    
    if (isMobile) {
        // إضافة class للجسم للإشارة إلى أننا على شاشة صغيرة
        document.body.classList.add('mobile-view');
        
        // إعادة رسم الجدول لضمان ظهور البطاقات
        setTimeout(() => {
            renderTable();
        }, 100);
    } else {
        document.body.classList.remove('mobile-view');
    }
}

// ✅ إضافة مستمع لتغير حجم الشاشة
window.addEventListener('resize', function() {
    clearTimeout(this.resizeTimer);
    this.resizeTimer = setTimeout(enhanceMobileCards, 250);
});

// ✅ استدعاء الدالة عند التحميل
window.addEventListener('load', enhanceMobileCards);
