// Initialize Lucide icons
lucide.createIcons();

// State
let employeesData = [];
let selectedEmployee = null;
let currentWorkbook = null;

// DOM Elements
const fileUploadMain = document.getElementById('file-upload-main');
const fileUploadSmall = document.getElementById('file-upload-small');
const downloadSampleBtn = document.getElementById('download-sample-btn');
const sheetSelector = document.getElementById('sheet-selector');
const emptyState = document.getElementById('empty-state');
const sidebar = document.getElementById('sidebar');
const noSelectionState = document.getElementById('no-selection-state');
const detailView = document.getElementById('detail-view');
const employeeList = document.getElementById('candidate-list'); // Keeping ID candidate-list in JS to avoid HTML rename issues if missed
const searchInput = document.getElementById('search-input');
const sidebarEmpty = document.getElementById('sidebar-empty');

// Detail View Elements
const detailHeaderInfo = document.getElementById('detail-header-info');
const detailInitials = document.getElementById('detail-initials');
const detailPrimarySkills = document.getElementById('detail-primary-skills');
const detailSecondarySkills = document.getElementById('detail-secondary-skills');
const detailCertifications = document.getElementById('detail-certifications');

const sectionPrimarySkills = document.getElementById('section-primary-skills');
const sectionSecondarySkills = document.getElementById('section-secondary-skills');
const sectionCertifications = document.getElementById('section-certifications');

// Mobile menu
const mobileMenuBtn = document.getElementById('mobile-menu-btn');
const mobileOverlay = document.getElementById('mobile-overlay');

// Event Listeners
fileUploadMain.addEventListener('change', handleFileUpload);
fileUploadSmall.addEventListener('change', handleFileUpload);
searchInput.addEventListener('input', handleSearch);
downloadSampleBtn?.addEventListener('click', downloadSampleTemplate);
sheetSelector?.addEventListener('change', handleSheetChange);

mobileMenuBtn?.addEventListener('click', toggleMobileMenu);
mobileOverlay?.addEventListener('click', toggleMobileMenu);

function toggleMobileMenu() {
    sidebar.classList.toggle('hidden');
    sidebar.classList.toggle('absolute');
    sidebar.classList.toggle('z-30');
    sidebar.classList.toggle('h-full');
    mobileOverlay.classList.toggle('hidden');
}

function downloadSampleTemplate() {
    const sampleData = [
        {
            "Employee Name": "John Doe",
            "Role": "Frontend Developer",
            "Email ID": "john.doe@example.com",
            "PODS": "Alpha Pod",
            "Reporting Manager": "Alice Johnson",
            "Engagement Type": "Full-Time",
            "Primary Skills": "JavaScript, React, Node.js",
            "Secondary Skills": "HTML, CSS, Git",
            "Certifications": "AWS Developer Associate"
        },
        {
            "Employee Name": "Jane Smith",
            "Role": "Backend Engineer",
            "Email ID": "jane.smith@example.com",
            "PODS": "Beta Pod",
            "Reporting Manager": "Bob Miller",
            "Engagement Type": "Contractor",
            "Primary Skills": "Python, Django, PostgreSQL",
            "Secondary Skills": "Docker, Linux, REST APIs",
            "Certifications": "Google Cloud Professional"
        }
    ];

    const worksheet = XLSX.utils.json_to_sheet(sampleData);
    
    // Set some column widths for better formatting
    worksheet['!cols'] = [
        { wch: 20 }, // Employee Name
        { wch: 20 }, // Role
        { wch: 25 }, // Email
        { wch: 15 }, // PODS
        { wch: 20 }, // Manager
        { wch: 15 }, // Engagement
        { wch: 30 }, // Primary Skills
        { wch: 30 }, // Secondary Skills
        { wch: 35 }  // Certifications
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Employees");

    // Generate file and trigger download
    XLSX.writeFile(workbook, "employee_template.xlsx");
}

// File Upload Handling
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        currentWorkbook = XLSX.read(data, { type: 'array' });
        
        if (sheetSelector) {
            sheetSelector.innerHTML = '';
            currentWorkbook.SheetNames.forEach(name => {
                const opt = document.createElement('option');
                opt.value = name;
                opt.textContent = name;
                sheetSelector.appendChild(opt);
            });
            
            if (currentWorkbook.SheetNames.length > 0) {
                sheetSelector.classList.remove('hidden');
            } else {
                sheetSelector.classList.add('hidden');
            }
        }

        // Load first sheet by default
        if (currentWorkbook.SheetNames.length > 0) {
            loadSheet(currentWorkbook.SheetNames[0]);
        }
    };
    reader.readAsArrayBuffer(file);
    
    // Reset inputs
    fileUploadMain.value = '';
    fileUploadSmall.value = '';
}

function handleSheetChange(event) {
    if (!currentWorkbook) return;
    loadSheet(event.target.value);
}

function loadSheet(sheetName) {
    const worksheet = currentWorkbook.Sheets[sheetName];
    const rawJson = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    processExcelData(rawJson);
}

function processExcelData(rawJson) {
    if (rawJson.length === 0) {
        alert("The selected sheet is empty.");
        return;
    }

    // Helper to find column key case-insensitively or loosely
    const findKey = (obj, ...targetStrs) => {
        const targets = targetStrs.map(t => t.toLowerCase().replace(/[^a-z0-9]/g, ''));
        return Object.keys(obj).find(k => {
            const normalizedKey = k.toLowerCase().replace(/[^a-z0-9]/g, '');
            return targets.includes(normalizedKey);
        });
    };

    employeesData = rawJson.map(row => {
        const nameKey = findKey(row, 'name', 'employeename', 'employee');
        const pSkillsKey = findKey(row, 'primaryskills', 'primaryskill', 'pskills', 'coreskills');
        const sSkillsKey = findKey(row, 'secondaryskills', 'secondaryskill', 'secondryskills', 'secondryskill', 'sskills', 'otherskills');
        const certsKey = findKey(row, 'certifications', 'certification', 'certs', 'certificates');
        const emailKey = findKey(row, 'email', 'emailid', 'mail');
        const roleKey = findKey(row, 'role', 'designation', 'title');
        const podsKey = findKey(row, 'pod', 'pods');
        const managerKey = findKey(row, 'reportingmanager', 'manager');
        const engagementKey = findKey(row, 'engagementtype', 'engagement');

        return {
            id: Math.random().toString(36).substr(2, 9),
            name: nameKey && row[nameKey] ? row[nameKey] : 'Unknown Employee',
            email: emailKey && row[emailKey] ? row[emailKey] : '',
            role: roleKey && row[roleKey] ? row[roleKey] : '',
            pods: podsKey && row[podsKey] ? row[podsKey] : '',
            manager: managerKey && row[managerKey] ? row[managerKey] : '',
            engagement: engagementKey && row[engagementKey] ? row[engagementKey] : '',
            primarySkills: pSkillsKey ? String(row[pSkillsKey]).split(/[,|;]+/).map(s => s.trim()).filter(Boolean) : [],
            secondarySkills: sSkillsKey ? String(row[sSkillsKey]).split(/[,|;]+/).map(s => s.trim()).filter(Boolean) : [],
            certifications: certsKey ? String(row[certsKey]).split(/[,|;]+/).map(s => s.trim()).filter(Boolean) : []
        };
    }).filter(e => e.name !== 'Unknown Employee' && e.name.trim() !== ''); // More strict filter

    if (employeesData.length > 0) {
        // UI Transition
        emptyState.classList.add('hidden');
        sidebar.classList.remove('hidden'); // Show on md+ screens
        sidebar.classList.add('md:flex');
        noSelectionState.classList.remove('hidden');
        detailView.classList.add('hidden'); // Hide detail view when loading new sheet until clicked
        
        renderSidebar(employeesData);
    } else {
        alert("No valid employee data found in this sheet.");
    }
}

function renderSidebar(data) {
    employeeList.innerHTML = '';
    
    if (data.length === 0) {
        sidebarEmpty.classList.remove('hidden');
    } else {
        sidebarEmpty.classList.add('hidden');
        
        data.forEach(employee => {
            const li = document.createElement('li');
            
            const btn = document.createElement('button');
            btn.className = 'w-full text-left px-4 py-3 rounded-lg hover:bg-slate-50 transition-colors border border-transparent focus:outline-none flex items-center gap-3';
            
            // Get initials
            const initials = getInitials(employee.name);

            btn.innerHTML = `
                <div class="w-10 h-10 rounded-full bg-brand-100 flex items-center justify-center text-brand-600 font-bold text-sm shrink-0">
                    ${initials}
                </div>
                <div class="overflow-hidden">
                    <div class="font-medium text-slate-900 truncate">${employee.name}</div>
                    <div class="text-xs text-slate-500 truncate">${employee.role || (employee.primarySkills.slice(0,2).join(', ') || 'Employee')}</div>
                </div>
            `;
            
            btn.addEventListener('click', () => {
                // Update active state
                document.querySelectorAll('#candidate-list button').forEach(b => {
                    b.classList.remove('bg-brand-50', 'border-brand-100');
                    b.classList.add('hover:bg-slate-50', 'border-transparent');
                });
                btn.classList.remove('hover:bg-slate-50', 'border-transparent');
                btn.classList.add('bg-brand-50', 'border-brand-100');
                
                renderEmployeeDetail(employee);

                // Close mobile menu if open
                if (window.innerWidth < 768) {
                    toggleMobileMenu();
                }
            });

            li.appendChild(btn);
            employeeList.appendChild(li);
        });
    }
}

function handleSearch(e) {
    const query = e.target.value.toLowerCase();
    const filtered = employeesData.filter(emp => emp.name.toLowerCase().includes(query));
    renderSidebar(filtered);
}

function renderEmployeeDetail(employee) {
    selectedEmployee = employee;
    
    // UI Transition
    noSelectionState.classList.add('hidden');
    detailView.classList.remove('hidden');
    
    // Header Info
    detailInitials.textContent = getInitials(employee.name);
    
    let headerHtml = `
        <div class="flex flex-col md:flex-row md:items-start justify-between gap-4 w-full">
            <div>
                <h2 class="text-3xl font-bold text-slate-800 mb-1">${employee.name}</h2>
                ${employee.role ? `<p class="text-lg text-slate-500 font-medium mb-2">${employee.role}</p>` : ''}
            </div>
            ${employee.pods ? `<div class="mt-1"><span class="inline-flex items-center px-4 py-1.5 rounded-full text-sm font-semibold bg-indigo-100 text-indigo-700 border border-indigo-200">${employee.pods}</span></div>` : ''}
        </div>
    `;

    const gridItems = [];
    if (employee.email) {
        gridItems.push(`
            <div class="flex items-center gap-2 text-sm text-slate-600">
                <i data-lucide="mail" class="w-4 h-4 text-slate-400"></i>
                <span class="truncate">${employee.email}</span>
            </div>
        `);
    }
    if (employee.manager) {
        gridItems.push(`
            <div class="flex items-center gap-2 text-sm text-slate-600">
                <i data-lucide="user" class="w-4 h-4 text-slate-400"></i>
                <span class="truncate">Manager: ${employee.manager}</span>
            </div>
        `);
    }
    if (employee.engagement) {
        gridItems.push(`
            <div class="flex items-center gap-2 text-sm text-slate-600">
                <i data-lucide="briefcase" class="w-4 h-4 text-slate-400"></i>
                <span class="truncate">${employee.engagement}</span>
            </div>
        `);
    }

    if (gridItems.length > 0) {
        headerHtml += `<div class="grid grid-cols-1 md:grid-cols-3 gap-4 mt-6 pt-6 border-t border-slate-100 w-full">
            ${gridItems.join('')}
        </div>`;
    }

    detailHeaderInfo.innerHTML = headerHtml;
    // Re-initialize icons inside the new dynamically added HTML
    lucide.createIcons({ root: detailHeaderInfo });
    
    // Primary Skills
    if (employee.primarySkills.length > 0) {
        if (sectionPrimarySkills) sectionPrimarySkills.classList.remove('hidden');
        detailPrimarySkills.innerHTML = '';
        employee.primarySkills.forEach(skill => {
            const span = document.createElement('span');
            span.className = 'inline-flex items-center px-3 py-1 rounded-full text-sm font-medium bg-brand-100 text-brand-700 border border-brand-200';
            span.textContent = skill;
            detailPrimarySkills.appendChild(span);
        });
    } else {
        if (sectionPrimarySkills) sectionPrimarySkills.classList.add('hidden');
    }

    // Secondary Skills
    if (employee.secondarySkills.length > 0) {
        if (sectionSecondarySkills) sectionSecondarySkills.classList.remove('hidden');
        detailSecondarySkills.innerHTML = '';
        employee.secondarySkills.forEach(skill => {
            const li = document.createElement('li');
            li.className = 'flex items-center gap-2 text-sm text-slate-700 bg-slate-50 px-3 py-2 rounded-lg border border-slate-100';
            li.innerHTML = `<i data-lucide="check-circle-2" class="w-4 h-4 text-slate-400"></i> ${skill}`;
            detailSecondarySkills.appendChild(li);
        });
        lucide.createIcons({ root: detailSecondarySkills });
    } else {
        if (sectionSecondarySkills) sectionSecondarySkills.classList.add('hidden');
    }

    // Certifications
    if (employee.certifications.length > 0) {
        if (sectionCertifications) sectionCertifications.classList.remove('hidden');
        detailCertifications.innerHTML = '';
        employee.certifications.forEach(cert => {
            const div = document.createElement('div');
            div.className = 'flex items-start gap-3 p-3 rounded-xl bg-slate-50 border border-slate-100';
            div.innerHTML = `
                <div class="mt-0.5">
                    <i data-lucide="shield-check" class="w-5 h-5 text-emerald-500"></i>
                </div>
                <div class="text-sm font-medium text-slate-800">${cert}</div>
            `;
            detailCertifications.appendChild(div);
        });
        lucide.createIcons({ root: detailCertifications });
    } else {
        if (sectionCertifications) sectionCertifications.classList.add('hidden');
    }
}

function getInitials(name) {
    if (!name || name.trim() === '') return 'E';
    return name
        .split(' ')
        .filter(n => n.length > 0)
        .map(n => n[0])
        .slice(0, 2)
        .join('')
        .toUpperCase();
}
