// Initialize Lucide icons
lucide.createIcons();

// State
let candidatesData = [];
let selectedCandidate = null;

// DOM Elements
const fileUploadMain = document.getElementById('file-upload-main');
const fileUploadSmall = document.getElementById('file-upload-small');
const downloadSampleBtn = document.getElementById('download-sample-btn');
const emptyState = document.getElementById('empty-state');
const sidebar = document.getElementById('sidebar');
const noSelectionState = document.getElementById('no-selection-state');
const detailView = document.getElementById('detail-view');
const candidateList = document.getElementById('candidate-list');
const searchInput = document.getElementById('search-input');
const sidebarEmpty = document.getElementById('sidebar-empty');

// Detail View Elements
const detailHeaderInfo = document.getElementById('detail-header-info');
const detailInitials = document.getElementById('detail-initials');
const detailPrimarySkills = document.getElementById('detail-primary-skills');
const detailSecondarySkills = document.getElementById('detail-secondary-skills');
const detailCertifications = document.getElementById('detail-certifications');

// Mobile menu
const mobileMenuBtn = document.getElementById('mobile-menu-btn');
const mobileOverlay = document.getElementById('mobile-overlay');

// Event Listeners
fileUploadMain.addEventListener('change', handleFileUpload);
fileUploadSmall.addEventListener('change', handleFileUpload);
searchInput.addEventListener('input', handleSearch);
downloadSampleBtn?.addEventListener('click', downloadSampleTemplate);

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
            "Primary Skills": "JavaScript, React, Node.js",
            "Secondary Skills": "HTML, CSS, Git",
            "Certifications": "AWS Developer Associate"
        },
        {
            "Employee Name": "Jane Smith",
            "Primary Skills": "Python, Django, PostgreSQL",
            "Secondary Skills": "Docker, Linux, REST APIs",
            "Certifications": "Google Cloud Professional"
        }
    ];

    const worksheet = XLSX.utils.json_to_sheet(sampleData);
    
    // Set some column widths for better formatting
    worksheet['!cols'] = [
        { wch: 20 }, // Employee Name
        { wch: 30 }, // Primary Skills
        { wch: 30 }, // Secondary Skills
        { wch: 35 }  // Certifications
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Candidates");

    // Generate file and trigger download
    XLSX.writeFile(workbook, "candidate_template.xlsx");
}

// File Upload Handling
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Assume first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to JSON
        const rawJson = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        // Process and normalize data
        processExcelData(rawJson);
    };
    reader.readAsArrayBuffer(file);
    
    // Reset inputs
    fileUploadMain.value = '';
    fileUploadSmall.value = '';
}

function processExcelData(rawJson) {
    if (rawJson.length === 0) {
        alert("The uploaded Excel file is empty.");
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

    candidatesData = rawJson.map(row => {
        const nameKey = findKey(row, 'name', 'candidatename', 'candidate', 'employeename');
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
            name: nameKey && row[nameKey] ? row[nameKey] : 'Unknown Candidate',
            email: emailKey && row[emailKey] ? row[emailKey] : '',
            role: roleKey && row[roleKey] ? row[roleKey] : '',
            pods: podsKey && row[podsKey] ? row[podsKey] : '',
            manager: managerKey && row[managerKey] ? row[managerKey] : '',
            engagement: engagementKey && row[engagementKey] ? row[engagementKey] : '',
            primarySkills: pSkillsKey ? String(row[pSkillsKey]).split(/[,|;]+/).map(s => s.trim()).filter(Boolean) : [],
            secondarySkills: sSkillsKey ? String(row[sSkillsKey]).split(/[,|;]+/).map(s => s.trim()).filter(Boolean) : [],
            certifications: certsKey ? String(row[certsKey]).split(/[,|;]+/).map(s => s.trim()).filter(Boolean) : []
        };
    }).filter(c => c.name !== 'Unknown Candidate' || c.primarySkills.length > 0); // basic filter

    if (candidatesData.length > 0) {
        // UI Transition
        emptyState.classList.add('hidden');
        sidebar.classList.remove('hidden'); // Show on md+ screens
        sidebar.classList.add('md:flex');
        noSelectionState.classList.remove('hidden');
        
        renderSidebar(candidatesData);
    } else {
        alert("No valid candidate data found in the file.");
    }
}

function renderSidebar(data) {
    candidateList.innerHTML = '';
    
    if (data.length === 0) {
        sidebarEmpty.classList.remove('hidden');
    } else {
        sidebarEmpty.classList.add('hidden');
        
        data.forEach(candidate => {
            const li = document.createElement('li');
            
            const btn = document.createElement('button');
            btn.className = 'w-full text-left px-4 py-3 rounded-lg hover:bg-slate-50 transition-colors border border-transparent focus:outline-none flex items-center gap-3';
            
            // Get initials
            const initials = getInitials(candidate.name);

            btn.innerHTML = `
                <div class="w-10 h-10 rounded-full bg-brand-100 flex items-center justify-center text-brand-600 font-bold text-sm shrink-0">
                    ${initials}
                </div>
                <div class="overflow-hidden">
                    <div class="font-medium text-slate-900 truncate">${candidate.name}</div>
                    <div class="text-xs text-slate-500 truncate">${candidate.primarySkills.slice(0,2).join(', ') || 'No primary skills'}</div>
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
                
                renderCandidateDetail(candidate);

                // Close mobile menu if open
                if (window.innerWidth < 768) {
                    toggleMobileMenu();
                }
            });

            li.appendChild(btn);
            candidateList.appendChild(li);
        });
    }
}

function handleSearch(e) {
    const query = e.target.value.toLowerCase();
    const filtered = candidatesData.filter(c => c.name.toLowerCase().includes(query));
    renderSidebar(filtered);
}

function renderCandidateDetail(candidate) {
    selectedCandidate = candidate;
    
    // UI Transition
    noSelectionState.classList.add('hidden');
    detailView.classList.remove('hidden');
    
    // Header Info
    detailInitials.textContent = getInitials(candidate.name);
    
    let headerHtml = `
        <div class="flex flex-col md:flex-row md:items-start justify-between gap-4 w-full">
            <div>
                <h2 class="text-3xl font-bold text-slate-800 mb-1">${candidate.name}</h2>
                ${candidate.role ? `<p class="text-lg text-slate-500 font-medium mb-2">${candidate.role}</p>` : ''}
            </div>
            ${candidate.pods ? `<div class="mt-1"><span class="inline-flex items-center px-4 py-1.5 rounded-full text-sm font-semibold bg-indigo-100 text-indigo-700 border border-indigo-200">${candidate.pods}</span></div>` : ''}
        </div>
    `;

    const gridItems = [];
    if (candidate.email) {
        gridItems.push(`
            <div class="flex items-center gap-2 text-sm text-slate-600">
                <i data-lucide="mail" class="w-4 h-4 text-slate-400"></i>
                <span class="truncate">${candidate.email}</span>
            </div>
        `);
    }
    if (candidate.manager) {
        gridItems.push(`
            <div class="flex items-center gap-2 text-sm text-slate-600">
                <i data-lucide="user" class="w-4 h-4 text-slate-400"></i>
                <span class="truncate">Manager: ${candidate.manager}</span>
            </div>
        `);
    }
    if (candidate.engagement) {
        gridItems.push(`
            <div class="flex items-center gap-2 text-sm text-slate-600">
                <i data-lucide="briefcase" class="w-4 h-4 text-slate-400"></i>
                <span class="truncate">${candidate.engagement}</span>
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
    detailPrimarySkills.innerHTML = '';
    if (candidate.primarySkills.length > 0) {
        candidate.primarySkills.forEach(skill => {
            const span = document.createElement('span');
            span.className = 'inline-flex items-center px-3 py-1 rounded-full text-sm font-medium bg-brand-100 text-brand-700 border border-brand-200';
            span.textContent = skill;
            detailPrimarySkills.appendChild(span);
        });
    } else {
        detailPrimarySkills.innerHTML = '<span class="text-slate-400 text-sm italic">No primary skills listed.</span>';
    }

    // Secondary Skills
    detailSecondarySkills.innerHTML = '';
    if (candidate.secondarySkills.length > 0) {
        candidate.secondarySkills.forEach(skill => {
            const li = document.createElement('li');
            li.className = 'flex items-center gap-2 text-sm text-slate-700 bg-slate-50 px-3 py-2 rounded-lg border border-slate-100';
            li.innerHTML = `<i data-lucide="check-circle-2" class="w-4 h-4 text-slate-400"></i> ${skill}`;
            detailSecondarySkills.appendChild(li);
        });
        lucide.createIcons({ root: detailSecondarySkills });
    } else {
        detailSecondarySkills.innerHTML = '<li class="text-slate-400 text-sm italic w-full">No secondary skills listed.</li>';
    }

    // Certifications
    detailCertifications.innerHTML = '';
    if (candidate.certifications.length > 0) {
        candidate.certifications.forEach(cert => {
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
        detailCertifications.innerHTML = '<div class="text-slate-400 text-sm italic">No certifications listed.</div>';
    }
}

function getInitials(name) {
    return name
        .split(' ')
        .map(n => n[0])
        .slice(0, 2)
        .join('')
        .toUpperCase();
}
