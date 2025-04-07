class OutputTable {
    constructor(outputData) {
        this._dimensionTitles = outputData.dimension_titles;
        this._dimensions = outputData.dimensions;
        this._value_titles = outputData.value_titles
        this._values = outputData.values.flat()
        this.outputSection = document.querySelector('#output');
        this.dataSection = document.querySelector('#cell-data');

        this.textColor = [
                        'red',
                        'orange',
                        'yellow',
                        'green',
                        'blue',
                        'indigo',
                        'violet',
                        'pink']
        this.months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

        this._buildTable()
    }

    _buildTable() {
        let dimTable = this._buildDimensions()
        let valTable = this._buildValues()
        // console.log(valTable)
        let tableArray = this._combineTables(dimTable, valTable);

        this._buildTableHTML(tableArray)
    }

    _combineTables(dimTable, valTable) {
        let tableArray = Array(20)

        for (let i = 0; i < tableArray.length; i++) {
            tableArray[i] = [...dimTable[i], ...valTable[i]]
        }

        console.log(tableArray)
        return tableArray
    }

    _buildDimensions() {
        let dimTable = [this._dimensionTitles, ...this._dimensions];
        // console.log(dimTable)
        return dimTable
    }

    _buildValues() {
        const rows = this._dimensions.length
        const cols = this._value_titles.length
        let valTable = Array.from({ length: rows }, () => new Array(cols).fill(null));
        
        for (let i = 0; i < rows; i++) {
            const cells = this._values.filter(cell => cell.row == i)
            for (let j = 0; j < cols; j++) {
                const cell = cells.filter(item => item.column == j)
                
                if (cell[0]) {
                    valTable[i][j] = cell[0];
                } 
            }
        }
        let dates = this._value_titles.map(str => new Date(str.replace('GMT','')))
        dates = dates.map(date => `${this.months[date.getMonth()]}-${date.getFullYear()}`)
        valTable = [dates, ...valTable]

        return valTable
    }

    _buildTableHTML(tableArray) {
        const tableElement = document.createElement('table');
        tableElement.classList.add('output-table');

        let tableHTML = ''
        for (let i = 0; i < tableArray.length; i++) {
            // let rowHTML;

            if (i == 0) {
                tableHTML += `<tr>${this._buildTableHeader(tableArray[i])}</tr>`
            } else {
                tableHTML += `<tr>${this._buildTableData(tableArray[i])}</tr>`
            }
        }

        // console.log(tableHTML)

        tableElement.innerHTML = tableHTML;
        this.outputSection.querySelector('#table-wrap').appendChild(tableElement)

        const clickableCells = tableElement.querySelectorAll('td[data-cell-clickable=True]');

        clickableCells.forEach(cell => {
            cell.addEventListener('click', (e) => {
                console.log(this._values)
                const cellID = e.target.attributes['data-cell-id'].value;
                console.log(typeof cellID)

                const cellData = this._values.find(aCell => aCell.cell_id == cellID);
                console.log(cellData)
                this.displayCellData(cellData)
            })
        })
    }

    _buildTableHeader(tableRow) {
        // console.log(tableRow)
        let headerHTML = ''
        for (const cell of tableRow) {
            // console.log(cell)
            headerHTML += `<th>${cell}</th>`
        }

        return headerHTML
    }

    _buildTableData(tableRow) {
        let dataHTML = ''
        for (const cell of tableRow) {
            if (cell) {
                if (typeof cell == 'object') {
                    dataHTML += `<td data-cell-id=${cell.cell_id} data-cell-selected=False data-cell-clickable=True>${cell.calc_value.toFixed(2)}</td>`
                } else {
                    dataHTML += `<td>${cell}</td>`
                }
            } else {
                dataHTML += `<td></td>`
            }
            
        }

        return dataHTML
    }

    displayCellData(cellData) {
        this.dataSection.innerHTML = ''
        console.log(cellData)
        const cellContainer = document.createElement('div')
        cellContainer.classList.add('cell-data-container');

        let cellContainerHTML = `
        <div class="cell-data-wrap">
            <div class="cell-data calc-value">
                <p class="formula-text">Value: ${cellData.calc_value.toFixed(2)}</p>
            </div>
            <div class="cell-data formula">
                <div class="formula-wrap">
                    <p class="formula-text">
                        ${this.buildDataFormula(cellData.cell_ref)}
                    </p>
                </div>
            </div>
            <div class="cell-data data-description">
                ${this.buildDataDescription(cellData.cell_ref)}
            </div>
        </div>`
        
        cellContainer.innerHTML = cellContainerHTML;
        this.dataSection.appendChild(cellContainer)

        const formulaSpans = cellContainer.querySelectorAll('.formula-text .formula-value');

        formulaSpans.forEach(span => {
            span.addEventListener('mouseenter', (e) => {
                console.log('entered')
                const spanID = e.target.attributes['data-span-id'].value;
                const cellData = this.dataSection.querySelectorAll('.data-description-formatter');
                const matchingCell = Array.from(cellData).find(cell => cell.attributes['data-description-id'].value == spanID)
                
                matchingCell.classList.toggle('highlight-data');
            })

            span.addEventListener('mouseleave', (e) => {
                console.log('exited')
                const spanID = e.target.attributes['data-span-id'].value;
                const cellData = this.dataSection.querySelectorAll('.data-description-formatter');
                const matchingCell = Array.from(cellData).find(cell => cell.attributes['data-description-id'].value == spanID)

                matchingCell.classList.toggle('highlight-data');
            })
        })
        console.log(formulaSpans)
    }

    buildDataFormula(cell_ref) {
        let formulaHTML = ''
        let count = 0
        cell_ref.forEach(cell => {
            if (count == 0) {
                formulaHTML += `= <span data-span-color="${this.textColor[count]}" data-span-id="${count}" class="formula-value">${cell.value}</span>`
            } else {
                formulaHTML += ` * <span data-span-color="${this.textColor[count]}" data-span-id="${count}" class="formula-value">${cell.value}</span>`
            }

            count++;
        })

        return formulaHTML
    }

    buildDataDescription(cell_ref) {
        let descriptionHTML = ''
        let count = 0
        cell_ref.forEach(cell => {
            descriptionHTML += `
            <div class="data-description-formatter" data-description-id="${count}">
                <div class="data-description-wrap">
                    <div class="data-key">
                        <p class="key-text">Table:</p>
                        <p class="key-text">Value:</p>
                        <p class="key-text">Worksheet:</p>
                        <p class="key-text">Address:</p>
                    </div>
                    <div class="data-value">
                        <p class="value-text">${cell.table}</p>
                        <p class="value-text"><span data-span-color="${this.textColor[count]}" class="formula-value">${cell.value}</span></p>
                        <p class="value-text">${cell.sheet}</p>
                        <p class="value-text">${cell.address}</p>
                    </div>
                </div>
            </div>`
            

            count++;
        })

        return descriptionHTML
    }
}

let output;

const imgSelect = document.querySelector('.custom-file-upload');

const profileForm = document.querySelector('#profile-settings-form');
const profileInput = document.querySelector('#upload-excel');

const fileName = document.querySelector('#file-upload-name');
const fileType = document.querySelector('#file-upload-type');

const profileButton = document.querySelector('#profile-submit');
const profileImage = document.querySelector('#profile-settings-image');
const imageUrl = document.querySelector('#profile-image-url');

const statusElement = document.querySelector('#profile-upload-status');
const initialsElement = document.querySelector('.user-profile-initials.settings-profile');

imgSelect.addEventListener('mousedown', () => {
    imgSelect.style.backgroundColor = '#0a2369';
});

imgSelect.addEventListener('mouseup', () => {
    imgSelect.style.backgroundColor = '#0a236900';
});

profileButton.addEventListener('mousedown', () => {
    imgSelect.style.backgroundColor = '#0a2369';
});

profileButton.addEventListener('mouseup', () => {
    imgSelect.style.backgroundColor = '#0a236900';
});

profileForm.onsubmit = (e) => {
    e.preventDefault();
}

profileInput.onchange = (event) => {
    statusElement.innerHTML = '';
    statusElement.opacity = 0;

    fileName.innerHTML = profileInput.files[0].name;
    fileType.innerHTML = profileInput.files[0].type;

    fileName.style.opacity = 1;
    fileType.style.opacity = 1;

    setTimeout(() => {
        // this.settings.preview_image(profileInput, profileImage, imageUrl, initialsElement);
        profileButton.style.display = 'inline-block';
        profileButton.style.opacity = 1;
    }, 400);
}

document.querySelector('#profile-submit').addEventListener('click', (e) => {
    let formData = new FormData(profileForm);
    // console.log([...formData])
    excelFileHandler(formData);
    profileInput.value = '';
    fileName.style.opacity = 0;
    fileType.style.opacity = 0;

    profileButton.style.opacity = 0;
    setTimeout(() => {
        fileName.innerHTML = '';
        fileType.innerHTML = '';

        profileButton.style.display = 'none';
    }, 400);
})

function excelFileHandler(formData) {
    fetch(`http://127.0.0.1:5000/tricentis/submit_form`, {
        method: 'POST',
        body: formData,
    })
        .then(response => {
            response.json()
        .then(data => {
            console.log(data)
            output = new OutputTable(data);
        })
    })
    .catch(error => console.error(error));
    
}