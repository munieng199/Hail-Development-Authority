// متغيرات عامة
let dashboardData = null;
let statusChart = null;
let departmentChart = null;
const currentDateSerial = 45931; // التاريخ الحالي: 2025-10-01 (مصحح بناءً على التاريخ الحالي)

// دالة تحويل تسلسلي Excel إلى تاريخ (YYYY-MM-DD)
function excelSerialToDate(serial) {
    if (isNaN(serial) || serial === '' || serial === 'مستمرة' || /L\d+|A\d+/.test(serial)) {
        return serial || '-'; // الاحتفاظ بالقيمة الأصلية إذا كانت غير عددية
    }
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    if (isNaN(date_info.getTime())) return serial || '-'; // إذا كان التحويل غير صالح
    const year = date_info.getFullYear();
    const month = String(date_info.getMonth() + 1).padStart(2, '0');
    const day = String(date_info.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`; // عرض كما في الملف الأصلي (ميلادي)
}

// عند تحميل الصفحة
document.addEventListener('DOMContentLoaded', function() {
    initializeDashboard();
});

// تهيئة لوحة المعلومات
function initializeDashboard() {
    const storedData = localStorage.getItem('excelData');
    
    if (storedData) {
        try {
            dashboardData = JSON.parse(storedData);
            dashboardData.tasks.forEach(task => {
                task['تاريخ  بدء المهمه'] = excelSerialToDate(task['تاريخ  بدء المهمه']);
                task['التاريخ المتوقع لانهاء المهمة'] = excelSerialToDate(task['التاريخ المتوقع لانهاء المهمة']);
                task['التاريخ الفعلي لانتهاء المهمة'] = excelSerialToDate(task['التاريخ الفعلي لانتهاء المهمة']);
                
                // حساب متأخر بدقة، ومعاملة "التسليم اليوم" كمتأخر
                let status = task['الحالة'] || '';
                const expectedSerial = parseFloat(task['التاريخ المتوقع لانهاء المهمة']);
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
                
                // نسبة التقدم تلقائيًا
                if (!task['نسبة التقدم']) {
                    task['نسبة التقدم'] = status.includes('مكتمل') ? 1 : status.includes('جاري') ? 0.5 : 0.25;
                }
            });
            dashboardData.summary = calculateSummary(dashboardData.tasks);
            loadDashboard();
        } catch (error) {
            console.error('خطأ في تحليل البيانات:', error);
            showNoDataState();
        }
    } else {
        showNoDataState();
    }
}

// تحميل لوحة المعلومات
function loadDashboard() {
    document.getElementById('loadingState').style.display = 'none';
    document.getElementById('dashboardContent').style.display = 'block';
    updateStatsCards();
    createCharts();
    updateInfoLists();
    updateTablePreview();
    updateProgressBar();
}

// تحديث البطاقات الإحصائية
function updateStatsCards() {
    const summary = dashboardData.summary;
    
    document.getElementById('totalTasks').textContent = summary.totalTasks;
    document.getElementById('completedTasks').textContent = summary.completedTasks;
    document.getElementById('completedPercentage').textContent = 
        Math.round((summary.completedTasks / summary.totalTasks) * 100) + '%';
    
    document.getElementById('delayedTasks').textContent = summary.delayedTasks;
    document.getElementById('delayedPercentage').textContent = 
        Math.round((summary.delayedTasks / summary.totalTasks) * 100) + '%';
    
    document.getElementById('inProgressTasks').textContent = summary.inProgressTasks;
    document.getElementById('inProgressPercentage').textContent = 
        Math.round((summary.inProgressTasks / summary.totalTasks) * 100) + '%';
    
    animateNumbers();
}

// تحريك الأرقام
function animateNumbers() {
    const numbers = document.querySelectorAll('.stats-number');
    
    numbers.forEach(element => {
        const target = parseInt(element.textContent);
        let current = 0;
        const increment = target / 50;
        
        const timer = setInterval(() => {
            current += increment;
            if (current >= target) {
                current = target;
                clearInterval(timer);
            }
            element.textContent = Math.floor(current);
        }, 20);
    });
}

// إنشاء الرسوم البيانية
function createCharts() {
    createStatusChart();
    createDepartmentChart();
}

// إنشاء رسم بياني لحالات المهام
function createStatusChart() {
    const ctx = document.getElementById('statusChart').getContext('2d');
    const summary = dashboardData.summary;
    
    statusChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['المكتملة', 'المتأخرة', 'قيد العمل'],
            datasets: [{
                data: [summary.completedTasks, summary.delayedTasks, summary.inProgressTasks],
                backgroundColor: [
                    '#4CAF50',
                    '#f44336',
                    '#ff9800'
                ],
                borderWidth: 2,
                borderColor: '#ffffff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        font: {
                            family: 'Cairo',
                            size: 12
                        },
                        padding: 20,
                        usePointStyle: true
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = Math.round((context.parsed / total) * 100);
                            return context.label + ': ' + context.parsed + ' (' + percentage + '%)';
                        }
                    },
                    titleFont: {
                        family: 'Cairo'
                    },
                    bodyFont: {
                        family: 'Cairo'
                    }
                }
            },
            animation: {
                animateRotate: true,
                duration: 2000
            }
        }
    });
}

// إنشاء رسم بياني للإدارات
function createDepartmentChart() {
    const ctx = document.getElementById('departmentChart').getContext('2d');
    
    const departmentCounts = {};
    dashboardData.tasks.forEach(task => {
        const dept = task['الإدارة'] || 'غير محدد';
        departmentCounts[dept] = (departmentCounts[dept] || 0) + 1;
    });
    
    const labels = Object.keys(departmentCounts);
    const data = Object.values(departmentCounts);
    
    departmentChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'عدد المهام',
                data: data,
                backgroundColor: '#006C35',
                borderColor: '#004d26',
                borderWidth: 1,
                borderRadius: 5
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    titleFont: {
                        family: 'Cairo'
                    },
                    bodyFont: {
                        family: 'Cairo'
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: {
                            family: 'Cairo'
                        }
                    }
                },
                x: {
                    ticks: {
                        font: {
                            family: 'Cairo'
                        },
                        maxRotation: 45
                    }
                }
            },
            animation: {
                duration: 2000,
                easing: 'easeInOutQuart'
            }
        }
    });
}

// تحديث قوائم المعلومات
function updateInfoLists() {
    const departmentsList = document.getElementById('departmentsList');
    departmentsList.innerHTML = '';
    
    dashboardData.summary.departments.forEach(dept => {
        const li = document.createElement('li');
        li.textContent = dept;
        departmentsList.appendChild(li);
    });
    
    const responsibleList = document.getElementById('responsibleList');
    responsibleList.innerHTML = '';
    
    dashboardData.summary.responsiblePersons.forEach(person => {
        const li = document.createElement('li');
        li.textContent = person;
        responsibleList.appendChild(li);
    });
}

// تحديث معاينة الجدول
function updateTablePreview() {
    const tbody = document.getElementById('tasksPreviewBody');
    tbody.innerHTML = '';
    
    const previewTasks = dashboardData.tasks.slice(0, 5);
    
    previewTasks.forEach(task => {
        const row = document.createElement('tr');
        
        const taskName = task['الموضوع/المهمة'].toLowerCase();
        if (taskName.includes('efqm') || taskName.includes('iso')) {
            row.classList.add('highlight-yellow');
        } else if (taskName.includes('خطة صرف')) {
            row.classList.add('highlight-orange');
        } else if (taskName.includes('27001') || taskName.includes('الأمن السيبراني')) {
            row.classList.add('highlight-blue');
        }
        
        const taskCell = document.createElement('td');
        taskCell.textContent = task['الموضوع/المهمة'] || '-';
        taskCell.style.maxWidth = '200px';
        taskCell.style.overflow = 'hidden';
        taskCell.style.textOverflow = 'ellipsis';
        taskCell.style.whiteSpace = 'nowrap';
        row.appendChild(taskCell);
        
        const deptCell = document.createElement('td');
        deptCell.textContent = task['الإدارة'] || '-';
        row.appendChild(deptCell);
        
        const responsibleCell = document.createElement('td');
        responsibleCell.textContent = task['المسؤول عن المهمه'] || '-';
        row.appendChild(responsibleCell);
        
        const startCell = document.createElement('td');
        startCell.textContent = task['تاريخ  بدء المهمه'] || '-';
        row.appendChild(startCell);
        
        const expectedCell = document.createElement('td');
        expectedCell.textContent = task['التاريخ المتوقع لانهاء المهمة'] || '-';
        row.appendChild(expectedCell);
        
        const statusCell = document.createElement('td');
        const statusBadge = document.createElement('span');
        const status = task['الحالة'] || '';
        
        statusBadge.textContent = status || 'غير محدد';
        statusBadge.className = 'status-badge';
        
        if (status.includes('مكتمل')) {
            statusBadge.classList.add('status-completed');
        } else if (status.includes('متأخر')) {
            statusBadge.classList.add('status-delayed');
        } else {
            statusBadge.classList.add('status-in-progress');
        }
        
        statusCell.appendChild(statusBadge);
        row.appendChild(statusCell);
        
        const progressCell = document.createElement('td');
        const progress = task['نسبة التقدم'] || 0;
        progressCell.textContent = (progress * 100).toFixed(0) + '%';
        row.appendChild(progressCell);
        
        tbody.appendChild(row);
    });
}

// تحديث شريط التقدم
function updateProgressBar() {
    const summary = dashboardData.summary;
    const completionRate = summary.completionRate;
    
    document.getElementById('overallProgress').textContent = completionRate + '%';
    
    setTimeout(() => {
        document.getElementById('progressFill').style.width = completionRate + '%';
    }, 500);
}

// عرض حالة عدم وجود بيانات
function showNoDataState() {
    document.getElementById('loadingState').style.display = 'none';
    document.getElementById('dashboardContent').style.display = 'none';
    document.getElementById('noDataState').style.display = 'block';
}

// العودة للصفحة الرئيسية
function goBack() {
    window.location.href = 'index.html';
}

// عرض الجدول الكامل
function showFullTable() {
    window.location.href = 'full-table.html';
}

// تنظيف الرسوم البيانية عند إغلاق الصفحة
window.addEventListener('beforeunload', function() {
    if (statusChart) {
        statusChart.destroy();
    }
    if (departmentChart) {
        departmentChart.destroy();
    }
});

// إعادة تحجيم الرسوم البيانية عند تغيير حجم النافذة
window.addEventListener('resize', function() {
    if (statusChart) {
        statusChart.resize();
    }
    if (departmentChart) {
        departmentChart.resize();
    }
});

// دالة حساب الملخص (تصحيح الإحصائيات بناءً على القيم المحددة)
// دالة حساب الملخص (تصحيح الإحصائيات)
function calculateSummary(tasks) {
    const summary = {
        totalTasks: 0, // سيتم حسابه من البيانات
        completedTasks: 0, // سيتم حسابه من البيانات
        delayedTasks: 0, // سيتم حسابه من البيانات
        inProgressTasks: 0, // سيتم حسابه من البيانات
        departments: new Set(),
        responsiblePersons: new Set()
    };
    
    // قائمة الإدارات المسموح بها فقط
    const allowedDepartments = [
        'الإدارة العامة للتميز المؤسسي',
        'إدارة الجودة الشاملة',
        'ادارة تميز الاعمال',
        'وحدة البحث والابتكار'
    ];
    
    // قائمة الأسماء المسموح بها فقط
    const allowedNames = [
        'ابراهيم البدر',
        'محمد الطواله',
        'علي حكمي',
        'عبداللطيف الهمشي',
        'تركي الباتع',
        'سعد البطي'
    ];
    
    tasks.forEach(task => {
        // حساب الإجمالي
        summary.totalTasks++;
        
        // حساب حالات المهام
        const status = task['الحالة'] || '';
        if (status.includes('مكتمل')) {
            summary.completedTasks++;
        } else if (status.includes('متأخر')) {
            summary.delayedTasks++;
        } else if (status.includes('جاري')) {
            summary.inProgressTasks++;
        }
        
        // جمع الإدارات مع التحقق من القائمة المسموح بها
        if (task['الإدارة']) {
            const deptName = task['الإدارة'].trim();
            if (allowedDepartments.some(allowed => 
                deptName.includes(allowed.split(' ')[0]) || 
                allowed.includes(deptName)
            )) {
                summary.departments.add(deptName);
            }
        }
        
        // جمع المسؤولين مع التحقق من القائمة المسموح بها
        if (task['المسؤول عن المهمه']) {
            const respName = task['المسؤول عن المهمه'].trim();
            
            // البحث عن تطابق مع أي اسم من القائمة المسموح بها
            const matchedName = allowedNames.find(allowedName => {
                // تقسيم الاسم المسموح به إلى أجزاء
                const nameParts = allowedName.split(' ');
                
                // التحقق من أن جميع أجزاء الاسم موجودة في اسم المسؤول
                return nameParts.every(part => 
                    respName.includes(part)
                );
            });
            
            // إذا وجد تطابق، أضف الاسم الأساسي (من القائمة المسموح بها)
            if (matchedName) {
                summary.responsiblePersons.add(matchedName);
            }
        }
    });
    
    // تحويل Sets إلى arrays وترتيبها
    summary.departments = Array.from(summary.departments).sort();
    summary.responsiblePersons = Array.from(summary.responsiblePersons).sort();
    summary.completionRate = summary.totalTasks > 0 ? Math.round((summary.completedTasks / summary.totalTasks) * 100) : 0;
    
    return summary;
}
// دالة فلتر المهام عند النقر على الكارد
function filterTasks(status) {
    if (status === 'total') {
        localStorage.removeItem('taskStatusFilter');
    } else {
        localStorage.setItem('taskStatusFilter', status);
    }
    window.location.href = 'full-table.html';
}
