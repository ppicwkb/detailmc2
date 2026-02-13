
        const API_KEY = "AIzaSyBhiKDVDH4fle5_EqAIaA05YjpxVMEBYZM";
        const SHEET_ID = "1jeVUjgFmTDVNsOGbKgb9Lglp8guMCTh-bNapK9owO8k";
        const SHEET_NAME = "DETAIL";
        
        let allData = [];
        let filteredData = [];
        let currentPage = 1;
        let itemsPerPage = 40;
        let totalPages = 1;
        let currentZoom = 100;
        let uniqueValues = {
            produk: new Set(),
            packing: new Set(),
            brand: new Set(),
            po: new Set(),
            kode: new Set()
        };

        // Load last modified date from Google Drive API
        async function loadLastModified() {
            try {
                const driveUrl = `https://www.googleapis.com/drive/v3/files/${SHEET_ID}?fields=modifiedTime,name&key=${API_KEY}`;
                const response = await fetch(driveUrl);
                const data = await response.json();
                
                if (data.modifiedTime) {
                    const modifiedDate = new Date(data.modifiedTime);
                    const formattedDate = modifiedDate.toLocaleString('id-ID', {
                        weekday: 'long',
                        year: 'numeric',
                        month: 'long',
                        day: 'numeric',
                        hour: '2-digit',
                        minute: '2-digit',
                        timeZone: 'Asia/Jakarta'
                    });
                    
                    document.getElementById('lastModified').innerHTML = `
                        <svg class="w-4 h-4 mr-1 text-green-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path>
                        </svg>
                        <span>Terakhir diubah: <strong>${formattedDate}</strong></span>
                    `;
                } else {
                    document.getElementById('lastModified').innerHTML = `
                        <svg class="w-4 h-4 mr-1 text-yellow-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v3.75m9-.75a9 9 0 11-18 0 9 9 0 0118 0zm-9 3.75h.008v.008H12v-.008z"></path>
                        </svg>
                        <span>Informasi waktu tidak tersedia</span>
                    `;
                }
            } catch (error) {
                console.error('Error loading last modified date:', error);
                document.getElementById('lastModified').innerHTML = `
                    <svg class="w-4 h-4 mr-1 text-red-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L3.732 16c-.77.833.192 2.5 1.732 2.5z"></path>
                    </svg>
                    <span>Gagal memuat informasi waktu</span>
                `;
            }
        }

        // Load data from Google Sheets
        async function loadData() {
            try {
                const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${SHEET_NAME}?key=${API_KEY}`;
                const response = await fetch(url);
                const data = await response.json();
                
                if (data.values && data.values.length > 1) {
                    const headers = data.values[0];
                    const rows = data.values.slice(1);
                    
                    // Find column indices
                    const colIndices = {
                        raw: headers.indexOf('RAW'),
                        plist: headers.indexOf('NO P.LIST'),
                        produkId: headers.indexOf('PRODUK ID'),
                        packing: headers.indexOf('PACKING'),
                        brand: headers.indexOf('BRAND'),
                        po: headers.indexOf('PO'),
                        jd: headers.indexOf('JD'),
                        kode: headers.indexOf('KODE'),
                        kodeAsal: headers.indexOf('KODE ASAL'),
                        cek: headers.indexOf('CEK'),
                        saldo: headers.indexOf('SALDO')
                    };
                    
                    allData = rows.map(row => {
                        const rawValue = row[colIndices.raw] || '';
                        const rawFirst4 = rawValue.substring(0, 3);
                        let rawDisplay = rawFirst4;
                        if (rawFirst4 !== 'GIM' && rawFirst4 !== 'CON') {
                            rawDisplay = 'WKB';
                        }
                        
                        const jd = row[colIndices.jd] || '';
                        const kode = row[colIndices.kode] || '';
                        const kodeJadi = jd && kode ? `${jd} ${kode}` : (jd || kode || '');
                        
                        const item = {
                            raw: rawDisplay,
                            plist: row[colIndices.plist] || '',
                            produkId: row[colIndices.produkId] || '',
                            packing: row[colIndices.packing] || '',
                            brand: row[colIndices.brand] || '',
                            po: row[colIndices.po] || '',
                            jd: jd,
                            kode: kode,
                            kodeAsal: row[colIndices.kodeAsal] || '',
                            cek: row[colIndices.cek] || '',
                            kodeJadi: kodeJadi,
                            saldo: row[colIndices.saldo] || ''
                        };
                        
                        // Collect unique values for suggestions
                        if (item.produkId) uniqueValues.produk.add(item.produkId);
                        if (item.packing) uniqueValues.packing.add(item.packing);
                        if (item.brand) uniqueValues.brand.add(item.brand);
                        if (item.po) uniqueValues.po.add(item.po);
                        if (item.kode) uniqueValues.kode.add(item.kode);
                        
                        return item;
                    });
                    
                    filteredData = [...allData];
                    displayData();
                    setupFilters();
                    setupPaginationControls();
                    hideLoading();
                } else {
                    throw new Error('No data found');
                }
            } catch (error) {
                console.error('Error loading data:', error);
                showError();
            }
        }

        function hideLoading() {
            document.getElementById('loadingIndicator').classList.add('hidden');
            document.getElementById('dataContainer').classList.remove('hidden');
        }

        function showError() {
            document.getElementById('loadingIndicator').classList.add('hidden');
            document.getElementById('errorMessage').classList.remove('hidden');
        }

        function displayData() {
            const tbody = document.getElementById('dataTableBody');
            tbody.innerHTML = '';
            
            // Calculate pagination
            totalPages = Math.ceil(filteredData.length / itemsPerPage);
            const startIndex = (currentPage - 1) * itemsPerPage;
            const endIndex = startIndex + itemsPerPage;
            const pageData = filteredData.slice(startIndex, endIndex);
            
            pageData.forEach((item, index) => {
                const row = document.createElement('tr');
                row.className = `hover:bg-blue-50 transition-colors duration-200 ${index % 2 === 0 ? 'bg-white' : 'bg-gray-50/50'}`;
                row.innerHTML = `
                    <td class="px-2 py-2 text-xs font-medium text-gray-900 whitespace-nowrap">${item.raw}</td>
                    <td class="px-2 py-2 text-xs text-gray-700 whitespace-nowrap">${item.plist}</td>
                    <td class="px-2 py-2 text-xs text-gray-700 whitespace-nowrap">${item.produkId}</td>
                    <td class="px-2 py-2 text-xs text-gray-700 whitespace-nowrap">${item.packing}</td>
                    <td class="px-2 py-2 text-xs text-gray-700 whitespace-nowrap">${item.brand}</td>
                    <td class="px-2 py-2 text-xs text-gray-700 whitespace-nowrap">${item.po}</td>
                    <td class="px-2 py-2 text-xs text-gray-700 whitespace-nowrap">${item.jd}</td>
                    <td class="px-2 py-2 text-xs text-gray-700 whitespace-nowrap">${item.kode}</td>
                    <td class="px-2 py-2 text-xs text-gray-700 whitespace-nowrap">${item.kodeAsal}</td>
                    <td class="px-2 py-2 text-xs text-gray-700 whitespace-nowrap">${item.cek}</td>
                    <td class="px-2 py-2 text-xs font-bold text-blue-600 whitespace-nowrap bg-blue-50/50 rounded-lg">${item.kodeJadi}</td>
                    <td class="px-2 py-2 text-xs font-semibold text-green-600 whitespace-nowrap">${item.saldo}</td>
                `;
                tbody.appendChild(row);
            });
            
            updatePaginationInfo();
            updatePaginationControls();
        }

        function updatePaginationInfo() {
            const startRecord = filteredData.length === 0 ? 0 : (currentPage - 1) * itemsPerPage + 1;
            const endRecord = Math.min(currentPage * itemsPerPage, filteredData.length);
            
            document.getElementById('recordCount').textContent = `Menampilkan ${startRecord}-${endRecord} dari ${filteredData.length} record (Total: ${allData.length})`;
            document.getElementById('paginationInfo').textContent = `Halaman ${currentPage} dari ${totalPages}`;
        }

        function updatePaginationControls() {
            const firstBtn = document.getElementById('firstPageBtn');
            const prevBtn = document.getElementById('prevPageBtn');
            const nextBtn = document.getElementById('nextPageBtn');
            const lastBtn = document.getElementById('lastPageBtn');
            const pageNumbers = document.getElementById('pageNumbers');
            
            // Enable/disable buttons
            firstBtn.disabled = currentPage === 1;
            prevBtn.disabled = currentPage === 1;
            nextBtn.disabled = currentPage === totalPages || totalPages === 0;
            lastBtn.disabled = currentPage === totalPages || totalPages === 0;
            
            // Generate page numbers
            pageNumbers.innerHTML = '';
            const maxVisiblePages = 5;
            let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
            let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
            
            if (endPage - startPage < maxVisiblePages - 1) {
                startPage = Math.max(1, endPage - maxVisiblePages + 1);
            }
            
            for (let i = startPage; i <= endPage; i++) {
                const pageBtn = document.createElement('button');
                pageBtn.className = `px-2 py-1 text-xs border rounded-lg transition-all duration-200 btn-hover shadow-sm ${i === currentPage ? 'bg-gradient-to-r from-blue-500 to-purple-600 text-white border-blue-500 shadow-lg' : 'bg-white text-gray-700 border-gray-200 hover:bg-blue-50 hover:border-blue-300'}`;
                pageBtn.textContent = i;
                pageBtn.onclick = () => goToPage(i);
                pageNumbers.appendChild(pageBtn);
            }
        }

        function goToPage(page) {
            if (page >= 1 && page <= totalPages) {
                currentPage = page;
                displayData();
            }
        }

        function setupFilters() {
            // Setup input filters with suggestions
            setupSuggestionFilter('produkFilter', 'produkSuggestions', 'produk');
            setupSuggestionFilter('packingFilter', 'packingSuggestions', 'packing');
            setupSuggestionFilter('brandFilter', 'brandSuggestions', 'brand');
            setupSuggestionFilter('poFilter', 'poSuggestions', 'po');
            setupSuggestionFilter('kodeFilter', 'kodeSuggestions', 'kode');
            
            // Setup dropdown filters
            setupDropdownFilter('produkDropdown', 'produk');
            setupDropdownFilter('packingDropdown', 'packing');
            setupDropdownFilter('brandDropdown', 'brand');
            setupDropdownFilter('poDropdown', 'po');
            setupDropdownFilter('kodeDropdown', 'kode');
            
            // Initial dropdown population
            updateAllDropdowns();
        }

        function setupSuggestionFilter(inputId, suggestionsId, filterType) {
            const input = document.getElementById(inputId);
            const suggestionsDiv = document.getElementById(suggestionsId);
            
            if (!input || !suggestionsDiv) {
                console.warn(`Elements not found: ${inputId} or ${suggestionsId}`);
                return;
            }
            
            input.addEventListener('input', function() {
                const value = this.value.toLowerCase();
                suggestionsDiv.innerHTML = '';
                
                if (value.length > 0) {
                    const availableOptions = getAvailableOptionsForFilter(filterType);
                    const filtered = availableOptions.filter(item => 
                        item.toLowerCase().includes(value)
                    ).slice(0, 10);
                    
                    if (filtered.length > 0) {
                        filtered.forEach(item => {
                            const div = document.createElement('div');
                            div.className = 'px-3 py-2 cursor-pointer suggestion-item text-sm hover:bg-gray-100 transition-colors';
                            div.textContent = item;
                            
                            // Use addEventListener instead of onclick for better event handling
                            div.addEventListener('click', function(e) {
                                e.preventDefault();
                                e.stopPropagation();
                                input.value = item;
                                suggestionsDiv.classList.add('hidden');
                                applyFilters();
                            });
                            
                            suggestionsDiv.appendChild(div);
                        });
                        suggestionsDiv.classList.remove('hidden');
                    } else {
                        suggestionsDiv.classList.add('hidden');
                    }
                } else {
                    suggestionsDiv.classList.add('hidden');
                }
                
                applyFilters();
            });
            
            // Handle focus events
            input.addEventListener('focus', function() {
                if (this.value.length > 0) {
                    // Re-trigger input event to show suggestions
                    this.dispatchEvent(new Event('input'));
                }
            });
            
            input.addEventListener('blur', function() {
                // Delay hiding to allow click events on suggestions
                setTimeout(() => {
                    if (suggestionsDiv) {
                        suggestionsDiv.classList.add('hidden');
                    }
                }, 200);
            });
            
            // Handle keyboard navigation
            input.addEventListener('keydown', function(e) {
                const suggestions = suggestionsDiv.querySelectorAll('.suggestion-item');
                if (suggestions.length === 0) return;
                
                let currentIndex = -1;
                suggestions.forEach((item, index) => {
                    if (item.classList.contains('bg-blue-100')) {
                        currentIndex = index;
                        item.classList.remove('bg-blue-100');
                    }
                });
                
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    currentIndex = Math.min(currentIndex + 1, suggestions.length - 1);
                    suggestions[currentIndex].classList.add('bg-blue-100');
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    currentIndex = Math.max(currentIndex - 1, 0);
                    suggestions[currentIndex].classList.add('bg-blue-100');
                } else if (e.key === 'Enter' && currentIndex >= 0) {
                    e.preventDefault();
                    suggestions[currentIndex].click();
                } else if (e.key === 'Escape') {
                    suggestionsDiv.classList.add('hidden');
                }
            });
        }

        function setupDropdownFilter(dropdownId, filterType) {
            const dropdown = document.getElementById(dropdownId);
            const inputMap = {
                'produkDropdown': 'produkFilter',
                'packingDropdown': 'packingFilter',
                'brandDropdown': 'brandFilter',
                'poDropdown': 'poFilter',
                'kodeDropdown': 'kodeFilter'
            };
            
            dropdown.addEventListener('change', function() {
                // Update corresponding input when dropdown is used
                const inputField = document.getElementById(inputMap[dropdownId]);
                if (inputField) {
                    inputField.value = this.value;
                    applyFilters();
                }
            });
            
            // Setup dropdown button click to show dropdown options
            const btnId = dropdownId.replace('Dropdown', 'DropdownBtn');
            const dropdownBtn = document.getElementById(btnId);
            const suggestionsId = dropdownId.replace('Dropdown', 'Suggestions');
            const suggestionsDiv = document.getElementById(suggestionsId);
            
            if (dropdownBtn && suggestionsDiv) {
                dropdownBtn.addEventListener('click', function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    // Hide all other suggestion divs
                    document.querySelectorAll('[id$="Suggestions"]').forEach(div => {
                        if (div.id !== suggestionsId) {
                            div.classList.add('hidden');
                        }
                    });
                    
                    // Toggle current suggestions
                    if (suggestionsDiv.classList.contains('hidden')) {
                        showDropdownOptions(suggestionsDiv, filterType, inputMap[dropdownId]);
                        suggestionsDiv.classList.remove('hidden');
                    } else {
                        suggestionsDiv.classList.add('hidden');
                    }
                });
                
                // Close dropdown when clicking outside - use event delegation
                const closeDropdownHandler = function(e) {
                    const filterContainer = document.getElementById(inputMap[dropdownId])?.closest('.group');
                    if (filterContainer && !filterContainer.contains(e.target)) {
                        suggestionsDiv.classList.add('hidden');
                    }
                };
                
                // Store handler reference for potential cleanup
                dropdownBtn._closeHandler = closeDropdownHandler;
                document.addEventListener('click', closeDropdownHandler);
            }
        }
        
        function showDropdownOptions(suggestionsDiv, filterType, inputId) {
            const availableOptions = getAvailableOptionsForFilter(filterType);
            const inputField = document.getElementById(inputId);
            
            if (!inputField || !suggestionsDiv) {
                console.warn(`Elements not found: ${inputId} or suggestionsDiv`);
                return;
            }
            
            suggestionsDiv.innerHTML = '';
            
            // Add "Semua" option
            const allDiv = document.createElement('div');
            allDiv.className = 'px-3 py-2 cursor-pointer suggestion-item text-sm text-gray-500 hover:bg-gray-100 transition-colors';
            allDiv.textContent = `Semua ${filterType.charAt(0).toUpperCase() + filterType.slice(1)}`;
            
            allDiv.addEventListener('click', function(e) {
                e.preventDefault();
                e.stopPropagation();
                inputField.value = '';
                suggestionsDiv.classList.add('hidden');
                applyFilters();
            });
            
            suggestionsDiv.appendChild(allDiv);
            
            // Add available options
            availableOptions.forEach(item => {
                const div = document.createElement('div');
                div.className = 'px-3 py-2 cursor-pointer suggestion-item text-sm hover:bg-gray-100 transition-colors';
                div.textContent = item;
                
                div.addEventListener('click', function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    inputField.value = item;
                    suggestionsDiv.classList.add('hidden');
                    applyFilters();
                });
                
                suggestionsDiv.appendChild(div);
            });
        }

        function getAvailableOptionsForFilter(filterType) {
            const currentFilters = getCurrentFilters();
            
            // Create a copy without the current filter to see what's available
            const filtersWithoutCurrent = { ...currentFilters };
            delete filtersWithoutCurrent[filterType];
            
            const availableData = getFilteredData(filtersWithoutCurrent);
            const fieldMap = {
                'produk': 'produkId',
                'packing': 'packing',
                'brand': 'brand',
                'po': 'po',
                'kode': 'kode'
            };
            
            return [...new Set(availableData.map(item => item[fieldMap[filterType]]).filter(val => val))].sort((a, b) => a.localeCompare(b, 'id', { numeric: true, sensitivity: 'base' }));
        }

        function getCurrentFilters() {
            const inputFilters = {
                produk: document.getElementById('produkFilter').value.toLowerCase(),
                packing: document.getElementById('packingFilter').value.toLowerCase(),
                brand: document.getElementById('brandFilter').value.toLowerCase(),
                po: document.getElementById('poFilter').value.toLowerCase(),
                kode: document.getElementById('kodeFilter').value.toLowerCase()
            };
            
            const dropdownFilters = {
                produk: document.getElementById('produkDropdown').value,
                packing: document.getElementById('packingDropdown').value,
                brand: document.getElementById('brandDropdown').value,
                po: document.getElementById('poDropdown').value,
                kode: document.getElementById('kodeDropdown').value
            };
            
            // Combine filters - dropdown takes precedence over input
            return {
                produk: dropdownFilters.produk || inputFilters.produk,
                packing: dropdownFilters.packing || inputFilters.packing,
                brand: dropdownFilters.brand || inputFilters.brand,
                po: dropdownFilters.po || inputFilters.po,
                kode: dropdownFilters.kode || inputFilters.kode
            };
        }

        function getFilteredData(filters) {
            return allData.filter(item => {
                const matchProduk = !filters.produk || 
                    (typeof filters.produk === 'string' && filters.produk.length > 0 ? 
                        item.produkId.toLowerCase().includes(filters.produk) : 
                        item.produkId === filters.produk);
                        
                const matchPacking = !filters.packing || 
                    (typeof filters.packing === 'string' && filters.packing.length > 0 ? 
                        item.packing.toLowerCase().includes(filters.packing) : 
                        item.packing === filters.packing);
                        
                const matchBrand = !filters.brand || 
                    (typeof filters.brand === 'string' && filters.brand.length > 0 ? 
                        item.brand.toLowerCase().includes(filters.brand) : 
                        item.brand === filters.brand);
                        
                const matchPo = !filters.po || 
                    (typeof filters.po === 'string' && filters.po.length > 0 ? 
                        item.po.toLowerCase().includes(filters.po) : 
                        item.po === filters.po);
                        
                const matchKode = !filters.kode || 
                    (typeof filters.kode === 'string' && filters.kode.length > 0 ? 
                        item.kode.toLowerCase().includes(filters.kode) : 
                        item.kode === filters.kode);
                
                return matchProduk && matchPacking && matchBrand && matchPo && matchKode;
            });
        }

        function updateAllDropdowns() {
            updateDropdown('produkDropdown', 'produk', 'Semua Produk');
            updateDropdown('packingDropdown', 'packing', 'Semua Packing');
            updateDropdown('brandDropdown', 'brand', 'Semua Brand');
            updateDropdown('poDropdown', 'po', 'Semua PO');
            updateDropdown('kodeDropdown', 'kode', 'Semua Kode');
        }

        function updateDropdown(dropdownId, filterType, defaultLabel) {
            const dropdown = document.getElementById(dropdownId);
            const currentValue = dropdown.value;
            const availableOptions = getAvailableOptionsForFilter(filterType);
            
            dropdown.innerHTML = `<option value="">${defaultLabel}</option>`;
            availableOptions.forEach(value => {
                const option = document.createElement('option');
                option.value = value;
                option.textContent = value;
                if (value === currentValue) option.selected = true;
                dropdown.appendChild(option);
            });
        }

        function applyFilters() {
            const filters = getCurrentFilters();
            filteredData = getFilteredData(filters);
            
            currentPage = 1;
            displayData();
            updateAllDropdowns();
        }

        function clearAllFilters() {
            // Clear input filters
            document.getElementById('produkFilter').value = '';
            document.getElementById('packingFilter').value = '';
            document.getElementById('brandFilter').value = '';
            document.getElementById('poFilter').value = '';
            document.getElementById('kodeFilter').value = '';
            
            // Clear dropdown filters
            document.getElementById('produkDropdown').value = '';
            document.getElementById('packingDropdown').value = '';
            document.getElementById('brandDropdown').value = '';
            document.getElementById('poDropdown').value = '';
            document.getElementById('kodeDropdown').value = '';
            
            filteredData = [...allData];
            currentPage = 1;
            displayData();
            updateAllDropdowns();
        }

        function setupPaginationControls() {
            // Pagination controls are handled by onclick attributes in HTML
            // This function is kept for compatibility but doesn't need to do anything
        }

        function exportToPDF() {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('l', 'mm', 'a4');
            
            doc.setFontSize(16);
            doc.text('Data Stok WKB', 14, 15);
            doc.setFontSize(10);
            doc.text(`Exported: ${new Date().toLocaleDateString('id-ID')}`, 14, 22);
            
            const headers = [['RAW', 'NO P.LIST', 'PRODUK ID', 'PACKING', 'BRAND', 'PO', 'JD', 'KODE', 'KODE ASAL', 'CEK', 'KODE JADI', 'SALDO']];
            const data = filteredData.map(item => [
                item.raw, item.plist, item.produkId, item.packing, item.brand,
                item.po, item.jd, item.kode, item.kodeAsal, item.cek, item.kodeJadi, item.saldo
            ]);
            
            doc.autoTable({
                head: headers,
                body: data,
                startY: 30,
                styles: { fontSize: 8 },
                headStyles: { fillColor: [66, 139, 202] }
            });
            
            doc.save('data-stok-wkb.pdf');
        }

        function exportToExcel() {
            const ws = XLSX.utils.json_to_sheet(filteredData.map(item => ({
                'RAW': item.raw,
                'NO P.LIST': item.plist,
                'PRODUK ID': item.produkId,
                'PACKING': item.packing,
                'BRAND': item.brand,
                'PO': item.po,
                'JD': item.jd,
                'KODE': item.kode,
                'KODE ASAL': item.kodeAsal,
                'CEK': item.cek,
                'KODE JADI': item.kodeJadi,
                'SALDO': item.saldo
            })));
            
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Data Stok wkb');
            XLSX.writeFile(wb, 'data-stok-wkb.xlsx');
        }

        function exportPivotTable() {
            // Create pivot table data with specified order: PO, Brand, Produk ID, Packing, Kode Jadi, Saldo
            const pivotData = {};
            
            // Group data by the specified hierarchy
            filteredData.forEach(item => {
                const po = item.po || 'Tidak Ada PO';
                const brand = item.brand || 'Tidak Ada Brand';
                const produkId = item.produkId || 'Tidak Ada Produk ID';
                const packing = item.packing || 'Tidak Ada Packing';
                const kodeJadi = item.kodeJadi || 'Tidak Ada Kode Jadi';
                const saldo = parseFloat(item.saldo) || 0;
                
                // Create nested structure
                if (!pivotData[po]) pivotData[po] = {};
                if (!pivotData[po][brand]) pivotData[po][brand] = {};
                if (!pivotData[po][brand][produkId]) pivotData[po][brand][produkId] = {};
                if (!pivotData[po][brand][produkId][packing]) pivotData[po][brand][produkId][packing] = {};
                if (!pivotData[po][brand][produkId][packing][kodeJadi]) {
                    pivotData[po][brand][produkId][packing][kodeJadi] = {
                        totalSaldo: 0,
                        count: 0
                    };
                }
                
                pivotData[po][brand][produkId][packing][kodeJadi].totalSaldo += saldo;
                pivotData[po][brand][produkId][packing][kodeJadi].count += 1;
            });
            
            // Convert to flat array for Excel export
            const flatData = [];
            
            Object.keys(pivotData).sort().forEach(po => {
                Object.keys(pivotData[po]).sort().forEach(brand => {
                    Object.keys(pivotData[po][brand]).sort().forEach(produkId => {
                        Object.keys(pivotData[po][brand][produkId]).sort().forEach(packing => {
                            Object.keys(pivotData[po][brand][produkId][packing]).sort().forEach(kodeJadi => {
                                const data = pivotData[po][brand][produkId][packing][kodeJadi];
                                flatData.push({
                                    'PO': po,
                                    'BRAND': brand,
                                    'PRODUK ID': produkId,
                                    'PACKING': packing,
                                    'KODE JADI': kodeJadi,
                                    'TOTAL SALDO': data.totalSaldo,
                                    'JUMLAH ITEM': data.count
                                });
                            });
                        });
                    });
                });
            });
            
            // Create summary data
            const summaryData = [];
            
            // Summary by PO
            const poSummary = {};
            Object.keys(pivotData).forEach(po => {
                let totalSaldo = 0;
                let totalItems = 0;
                
                Object.keys(pivotData[po]).forEach(brand => {
                    Object.keys(pivotData[po][brand]).forEach(produkId => {
                        Object.keys(pivotData[po][brand][produkId]).forEach(packing => {
                            Object.keys(pivotData[po][brand][produkId][packing]).forEach(kodeJadi => {
                                const data = pivotData[po][brand][produkId][packing][kodeJadi];
                                totalSaldo += data.totalSaldo;
                                totalItems += data.count;
                            });
                        });
                    });
                });
                
                poSummary[po] = { totalSaldo, totalItems };
            });
            
            Object.keys(poSummary).sort().forEach(po => {
                summaryData.push({
                    'KATEGORI': 'PO',
                    'NAMA': po,
                    'TOTAL SALDO': poSummary[po].totalSaldo,
                    'JUMLAH ITEM': poSummary[po].totalItems
                });
            });
            
            // Summary by Brand
            const brandSummary = {};
            Object.keys(pivotData).forEach(po => {
                Object.keys(pivotData[po]).forEach(brand => {
                    if (!brandSummary[brand]) brandSummary[brand] = { totalSaldo: 0, totalItems: 0 };
                    
                    Object.keys(pivotData[po][brand]).forEach(produkId => {
                        Object.keys(pivotData[po][brand][produkId]).forEach(packing => {
                            Object.keys(pivotData[po][brand][produkId][packing]).forEach(kodeJadi => {
                                const data = pivotData[po][brand][produkId][packing][kodeJadi];
                                brandSummary[brand].totalSaldo += data.totalSaldo;
                                brandSummary[brand].totalItems += data.count;
                            });
                        });
                    });
                });
            });
            
            Object.keys(brandSummary).sort().forEach(brand => {
                summaryData.push({
                    'KATEGORI': 'BRAND',
                    'NAMA': brand,
                    'TOTAL SALDO': brandSummary[brand].totalSaldo,
                    'JUMLAH ITEM': brandSummary[brand].totalItems
                });
            });
            
            // Create workbook with multiple sheets
            const wb = XLSX.utils.book_new();
            
            // Main pivot table sheet
            const ws1 = XLSX.utils.json_to_sheet(flatData);
            XLSX.utils.book_append_sheet(wb, ws1, 'Pivot Table Detail');
            
            // Summary sheet
            const ws2 = XLSX.utils.json_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(wb, ws2, 'Summary');
            
            // Original data sheet (filtered)
            const ws3 = XLSX.utils.json_to_sheet(filteredData.map(item => ({
                'RAW': item.raw,
                'NO P.LIST': item.plist,
                'PRODUK ID': item.produkId,
                'PACKING': item.packing,
                'BRAND': item.brand,
                'PO': item.po,
                'JD': item.jd,
                'KODE': item.kode,
                'KODE ASAL': item.kodeAsal,
                'CEK': item.cek,
                'KODE JADI': item.kodeJadi,
                'SALDO': item.saldo
            })));
            XLSX.utils.book_append_sheet(wb, ws3, 'Data Asli');
            
            // Save file
            const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
            XLSX.writeFile(wb, `pivot-table-stok-${timestamp}.xlsx`);
        }

        // Zoom functionality
        function adjustZoom(change) {
            const newZoom = Math.max(50, Math.min(200, currentZoom + change));
            if (newZoom !== currentZoom) {
                currentZoom = newZoom;
                applyZoom();
                updateZoomDisplay();
            }
        }

        function resetZoom() {
            currentZoom = 100;
            applyZoom();
            updateZoomDisplay();
        }

        function applyZoom() {
            const dataContainer = document.getElementById('dataContainer');
            const table = dataContainer.querySelector('table');
            
            // Apply zoom to the table
            table.style.transform = `scale(${currentZoom / 100})`;
            table.style.transformOrigin = 'top left';
            
            // Adjust container to accommodate scaled content
            const tableContainer = table.parentElement;
            tableContainer.style.width = `${100 * (100 / currentZoom)}%`;
            tableContainer.style.height = currentZoom < 100 ? 'auto' : `${100 * (currentZoom / 100)}%`;
            
            // Update button states
            const zoomOutBtn = document.getElementById('zoomOutBtn');
            const zoomInBtn = document.getElementById('zoomInBtn');
            
            zoomOutBtn.disabled = currentZoom <= 30;
            zoomInBtn.disabled = currentZoom >= 200;
            
            if (currentZoom <= 30) {
                zoomOutBtn.classList.add('opacity-50', 'cursor-not-allowed');
            } else {
                zoomOutBtn.classList.remove('opacity-50', 'cursor-not-allowed');
            }
            
            if (currentZoom >= 200) {
                zoomInBtn.classList.add('opacity-50', 'cursor-not-allowed');
            } else {
                zoomInBtn.classList.remove('opacity-50', 'cursor-not-allowed');
            }
        }

        function updateZoomDisplay() {
            document.getElementById('zoomLevel').textContent = `${currentZoom}%`;
        }

        // Keyboard shortcuts for zoom
        document.addEventListener('keydown', function(e) {
            if (e.ctrlKey || e.metaKey) {
                if (e.key === '=' || e.key === '+') {
                    e.preventDefault();
                    adjustZoom(10);
                } else if (e.key === '-') {
                    e.preventDefault();
                    adjustZoom(-10);
                } else if (e.key === '0') {
                    e.preventDefault();
                    resetZoom();
                }
            }
        });

        // Mouse wheel zoom (when holding Ctrl)
        document.addEventListener('wheel', function(e) {
            if (e.ctrlKey || e.metaKey) {
                e.preventDefault();
                const delta = e.deltaY > 0 ? -10 : 10;
                adjustZoom(delta);
            }
        }, { passive: false });

        // Touch zoom functionality
        let touchStartDistance = 0;
        let touchStartZoom = 100;
        let isZooming = false;

        function getTouchDistance(touches) {
            const dx = touches[0].clientX - touches[1].clientX;
            const dy = touches[0].clientY - touches[1].clientY;
            return Math.sqrt(dx * dx + dy * dy);
        }

        // Add touch event listeners to the data container
        const dataContainer = document.getElementById('dataContainer');
        
        dataContainer.addEventListener('touchstart', function(e) {
            if (e.touches.length === 2) {
                e.preventDefault();
                isZooming = true;
                touchStartDistance = getTouchDistance(e.touches);
                touchStartZoom = currentZoom;
            }
        }, { passive: false });

        dataContainer.addEventListener('touchmove', function(e) {
            if (e.touches.length === 2 && isZooming) {
                e.preventDefault();
                const currentDistance = getTouchDistance(e.touches);
                const scale = currentDistance / touchStartDistance;
                const newZoom = Math.max(50, Math.min(200, touchStartZoom * scale));
                
                if (Math.abs(newZoom - currentZoom) > 2) {
                    currentZoom = Math.round(newZoom);
                    applyZoom();
                    updateZoomDisplay();
                }
            }
        }, { passive: false });

        dataContainer.addEventListener('touchend', function(e) {
            if (e.touches.length < 2) {
                isZooming = false;
            }
        });

        // Prevent default zoom behavior on the data container
        dataContainer.addEventListener('gesturestart', function(e) {
            e.preventDefault();
        });

        dataContainer.addEventListener('gesturechange', function(e) {
            e.preventDefault();
        });

        dataContainer.addEventListener('gestureend', function(e) {
            e.preventDefault();
        });

        
        
        // Initialize the application
        async function init() {
            await loadLastModified();
            await loadData();
        }


        init();




