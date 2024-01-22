class DateUtilitaire {
    // Méthode statique pour obtenir la date actuelle formatée
    static dateFormatee() {
        const dateActuelle = new Date();

        const options = {
            day: 'numeric',
            month: 'numeric',
            year: 'numeric'
        };
        const dateFormatee = dateActuelle.toLocaleString('fr-FR', options);

        return dateFormatee;
    }

    // Méthode statique pour obtenir l'heure actuelle formatée
    static heureFormatee() {
        const heureActuelle = new Date();

        const options = {
            hour: 'numeric',
            minute: 'numeric',
            hour12: false
        };
        const heureFormatee = heureActuelle.toLocaleString('fr-FR', options);

        return heureFormatee;
    }
    static getWeekNumber(date) {
        var oneJan = new Date(date.getFullYear(),0,1);
        var numberOfDays = Math.floor((date - oneJan) / (24 * 60 * 60 * 1000));
        var result = Math.ceil((date.getDay() + 1 + numberOfDays) / 7);
        return result - 1;
    }
}

async function generateQRCodeAsync(text, size=200) {
    const qrDataURL = `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(text)}&size=${size}x${size}`;
    const qrImageDataURL = await getDataUri(qrDataURL);
    return qrImageDataURL;
}

jsPDF.API.addQRCodeToPage = async function addQRCodeToPage(qrCodeData, x, y, size) {
    const qrImageDataURL = await generateQRCodeAsync(qrCodeData, size);
    this.addImage(qrImageDataURL, 'PNG', x, y, size, size);
};

jsPDF.API.addUrlImageToPage = async function addUrlImageToPage(url, x, y, height, weight) {
    const urlImageData = await getDataUri(url);
    this.addImage(urlImageData, 'PNG', x, y, height, weight)
};

function getDataUri(url) {
    return new Promise(resolve=>{
        var image = new Image();
        image.setAttribute('crossOrigin', 'anonymous');
        //getting images from external domain

        image.onload = function() {
            var canvas = document.createElement('canvas');
            canvas.width = this.naturalWidth;
            canvas.height = this.naturalHeight;

            //next three lines for white background in case png has a transparent background
            var ctx = canvas.getContext('2d');
            ctx.fillStyle = '#fff';
            /// set white fill style
            ctx.fillRect(0, 0, canvas.width, canvas.height);

            canvas.getContext('2d').drawImage(this, 0, 0);

            resolve(canvas.toDataURL('image/jpeg'));
        }
        ;

        image.src = url;
    }
    )
}

class GenerationClass {
    constructor() {
        this.fileName = 'fileName';
        this.lastModified = new Date();
        this.weekGenerated = DateUtilitaire.getWeekNumber(this.lastModified);
    }
}



jsPDF.API.newCenteredText = function nouveauTexteCenter(font, format, size, text, y) {
    if (text == null || "") return;
    this.setFont(font, format);
    this.setFontSize(size);
    const largeurtext = this.getTextWidth(text);
    const xtext = (297 - largeurtext) / 2;
    // Centrer horizontalement
    this.text(text, xtext, y);
};

function addOptionsToDropdown(dropdown, options, storedValue) {
    // Effacer les anciennes options
    dropdown.innerHTML = '';

    // Ajouter une option par défaut
    const defaultOption = document.createElement('option');
    defaultOption.value = null;
    defaultOption.text = '👀 Choisir...';
    dropdown.add(defaultOption);

    // Ajouter les options du tableau Excel
    options.forEach(option => {
        const newOption = document.createElement('option');
        newOption.value = option;
        newOption.text = option;
        dropdown.add(newOption);
    });

    // Sélectionner la valeur stockée si elle existe
    dropdown.value = options.includes(storedValue) ? storedValue : defaultOption.value;
}

function generateDropdowns(sheetData, dropdownIds) {
    dropdownIds.forEach(id => {
        const dropdown = document.getElementById(id);
        const storedValue = getStoredValue(id + 'Choice');
        addOptionsToDropdown(dropdown, Object.keys(sheetData[0]), storedValue);
    });
};

function resetDropdowns(dropdownIds) {
    dropdownIds.forEach(id => {
        const dropdown = document.getElementById(id);
        dropdown.innerHTML = '';
        // Réinitialiser la valeur stockée
        setStoredValue(id + 'Choice', '');
    });
};

function saveDropdowns(dropdownIds) {
    dropdownIds.forEach(id => {
        const dropdown = document.getElementById(id);
        const selectedValue = dropdown.value;
        setStoredValue(id + 'Choice', selectedValue);
    });
}


function getValeurParNom(row, nom) {
    if (nom in row) {
        return row[nom];
    }
    return null;
    // Si la propriété n'est pas trouvée
};


function splitTextIntoLines(text, maxWidth, fontSize, font) {

    const lines = [];

    if (text == null) {
        console.error("Text is null");
        return lines;
    }

    const words = text.split(' ');
    let currentLine = '';

    const doc = new jsPDF();
    doc.setFont(font, 'normal');
    doc.setFontSize(fontSize);

    for (const word of words) {
        const currentWidth = doc.getTextWidth(currentLine + ' ' + word);
        if (currentWidth <= maxWidth) {
            if (currentLine !== '') {
                currentLine += ' ';
            }
            currentLine += word;
        } else {
            lines.push(currentLine);
            currentLine = word;
        }
    }

    if (currentLine !== '') {
        lines.push(currentLine);
    }

    return lines;
};

jsPDF.API.addBorders = function (marge, largeurPage, hauteurPage) {

    //#region CADRE

    this.setDrawColor(0);
    // Couleur du trait : noir
    this.setLineWidth(0.7);
    // Épaisseur du trait en points (1 point = 0.3528 mm)
    this.rect(marge + 1, marge + 1, largeurPage + 1.75 - (2 * marge), hauteurPage + 1.75 - (2 * marge), 'F');
    // Position x, y, largeur, hauteur
    this.setFillColor(255, 255, 255);
    // Couleur du fond : noir
    this.setLineWidth(0.75);
    // Épaisseur du trait en points (1 point = 0.3528 mm)
    this.rect(marge, marge, largeurPage - (2 * marge), hauteurPage - (2 * marge), 'DF');
    // Position x, y, largeur, hauteur
    this.setDrawColor(0);
    // Couleur du trait : noir
    this.setLineWidth(0.25);
    // Épaisseur du trait en points (1 point = 0.3528 mm)
    this.rect(marge + 2, marge + 2, largeurPage - 4 - (2 * marge), hauteurPage - 4 - (2 * marge), 'D');
    // Position x, y, largeur, hauteur

    //#endregion
};

// Fonction pour récupérer une valeur depuis le stockage local
function getStoredValue(key) {
    return localStorage.getItem(key);
}

// Fonction pour stocker une valeur dans le stockage local
function setStoredValue(key, value) {
    console.error(key)
    localStorage.setItem(key, value);
}