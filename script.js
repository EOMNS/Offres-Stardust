const qrCodeLink = 'https://candidat.pole-emploi.fr/offres/recherche/detail/';


class DateUtilitaire {
  // Méthode statique pour obtenir la date actuelle formatée
  static dateFormatee() {
    const dateActuelle = new Date();
    
    const options = { day: 'numeric', month: 'numeric', year: 'numeric' };
    const dateFormatee = dateActuelle.toLocaleString('fr-FR', options);

    return dateFormatee;
  }

  // Méthode statique pour obtenir l'heure actuelle formatée
  static heureFormatee() {
    const heureActuelle = new Date();
    
    const options = { hour: 'numeric', minute: 'numeric', hour12: false };
    const heureFormatee = heureActuelle.toLocaleString('fr-FR', options);

    return heureFormatee;
  }
  static getWeekNumber(date) {
    var oneJan = new Date(date.getFullYear(), 0, 1);
    var numberOfDays = Math.floor((date - oneJan) / (24 * 60 * 60 * 1000));
    var result = Math.ceil((date.getDay() + 1 + numberOfDays) / 7);
    return result;
  }
}

async function generateQRCodeAsync(text, size = 200) {
  const qrDataURL = `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(text)}&size=${size}x${size}`;
  const qrImageDataURL = await getDataUri(qrDataURL);
  return qrImageDataURL;
}

async function addQRCodeToPage(doc, qrCodeData, x, y, size) {
  const qrImageDataURL = await generateQRCodeAsync(qrCodeData, size);
  doc.addImage(qrImageDataURL, 'PNG', x, y, size, size);
}

async function addUrlImageToPage(doc, url, x, y, height, weight)
{
  const urlImageData = await getDataUri(url);
  doc.addImage(urlImageData, 'PNG', x, y, height, weight)
}

function getDataUri(url)
{
    return new Promise(resolve => {
        var image = new Image();
        image.setAttribute('crossOrigin', 'anonymous'); //getting images from external domain

        image.onload = function () {
            var canvas = document.createElement('canvas');
            canvas.width = this.naturalWidth;
            canvas.height = this.naturalHeight; 

            //next three lines for white background in case png has a transparent background
            var ctx = canvas.getContext('2d');
            ctx.fillStyle = '#fff';  /// set white fill style
            ctx.fillRect(0, 0, canvas.width, canvas.height);

            canvas.getContext('2d').drawImage(this, 0, 0);

            resolve(canvas.toDataURL('image/jpeg'));
        };

        image.src = url;
    })
}

class GenerationClass
{
  constructor()
  {
    this.fileName = 'fileName';
    this.lastModified = new Date();
    this.weekGenerated = DateUtilitaire.getWeekNumber(this.lastModified);
  }
}

var generation = new GenerationClass();


let sheetData

class Offre {
  constructor() {
      this.titre = 'titre';
      this.lieu = 'lieu';
      this.numero = 'numerooffre';
      this.type = 'typeoffre';
      this.duree = 'dureeoffre';
  }
}

function nouveauTexteCenter(doc, font, format, size, text, y)
{
  doc.setFont(font, format);
  doc.setFontSize(size);
  const largeurtext = doc.getTextWidth(text);
  const xtext = (297 - largeurtext) / 2; // Centrer horizontalement
  doc.text(text, xtext, y);
}

document.getElementById('inputFile').addEventListener('click', function(event)
{
  const cadrePDF = document.getElementById('cadrePDF');
  cadrePDF.innerHTML = '';
});

document.getElementById('inputFile').addEventListener('change', function(event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {
      const fileData = new Uint8Array(e.target.result);
      const workbook = XLSX.read(fileData, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      sheetData = XLSX.utils.sheet_to_json(worksheet);

      for (let i = 0; i < sheetData.length; i++) {
          const row = sheetData[i];
          console.log(row);
      }

      generation.fileName = file.name;
      generation.lastModified = new Date(file.lastModified);
  }

  reader.readAsArrayBuffer(file);
});

function getValeurParNom(row, nom) {
  if (nom in row) {
    return row[nom];
  }
  return null; // Si la propriété n'est pas trouvée
}

function splitTextIntoLines(text, maxWidth, fontSize, font) {
  const words = text.split(' ');
  let currentLine = '';
  const lines = [];

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
}


async function generatePDF() {

    
  const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
  });
  const largeurPage = 297; // Largeur de la page A4 en mm
  const hauteurPage = 210; // Hauteur de la page A4 en mm
  const marge = 10; // Marge entre le cadre et le contenu en mm

  const DateNow = new Date();


  const dateFormatee = DateUtilitaire.dateFormatee();
  const heureFormatee = DateUtilitaire.heureFormatee();
  await addUrlImageToPage(doc, 'https://eomns.github.io/Offres-Stardust/assets/icon.png', 0, largeurPage + (largeurPage/2), 200, 200)
  doc.setDrawColor(0); // Couleur du trait : noir
  doc.setLineWidth(0.7); // Épaisseur du trait en points (1 point = 0.3528 mm)
  doc.line(largeurPage/2, hauteurPage, largeurPage/2, -hauteurPage, 'S'); // Position x, y, largeur, hauteur

  doc.setFont('Arial', 'bold');
  doc.setFontSize(25)
  let text = `Offres de la semaine n°${generation.weekGenerated}`;
  doc.text(text, (largeurPage + (largeurPage/2) - doc.getTextWidth(text)) / 2, 105)
  doc.setFont('Arial', 'normal');
  doc.setFontSize(18)
  text = `Généré le ${dateFormatee} à ${heureFormatee}`;
  doc.text(text, (largeurPage + (largeurPage/2) - doc.getTextWidth(text)) / 2, 120)
  
  const lignesTitre = splitTextIntoLines(`d'après le tableau ${generation.fileName} modifié le ${DateUtilitaire.dateFormatee(generation.lastModified)}`, largeurPage/2 - (2 * marge+13)-10, 13, 'Arial');
  let yTitre = 160;
  for (const line of lignesTitre) {
    doc.setFont('Arial', 'normal');
    doc.setFontSize(13)
    const largeurTitre = doc.getTextWidth(line);
    const xTitre = (largeurPage + (largeurPage/2) - largeurTitre) / 2; // Centrer horizontalement
    doc.text(line, xTitre, yTitre);
    yTitre += 5; // Augmenter l'espacement entre les lignes
  }

  doc.addPage();

  const offre = new Offre();

  for (let i = 0; i < sheetData.length; i++) {
    let checktitle = getValeurParNom(sheetData[i], "Intitulé du poste")
    console.log(checktitle)
    if (checktitle) {
      offre.titre = checktitle
    } else {
      offre.titre = getValeurParNom(sheetData[i], "Appellation ROME")
      console.log(offre.titre)
    }

    offre.lieu = getValeurParNom(sheetData[i], "Lieu de travail")
    offre.numero = getValeurParNom(sheetData[i], "N° Offre")
    offre.type = getValeurParNom(sheetData[i], "Type contrat")
    offre.duree = getValeurParNom(sheetData[i], "Durée contrat")

    //#region CADRE

    doc.setDrawColor(0); // Couleur du trait : noir
    doc.setLineWidth(0.7); // Épaisseur du trait en points (1 point = 0.3528 mm)
    doc.rect(marge+1, marge+1, largeurPage+1.75 - (2 * marge), hauteurPage+1.75 - (2 * marge), 'F'); // Position x, y, largeur, hauteur
    doc.setFillColor(255,255,255); // Couleur du fond : noir
    doc.setLineWidth(0.75); // Épaisseur du trait en points (1 point = 0.3528 mm)
    doc.rect(marge, marge, largeurPage - (2 * marge), hauteurPage - (2 * marge), 'DF'); // Position x, y, largeur, hauteur
    doc.setDrawColor(0); // Couleur du trait : noir
    doc.setLineWidth(0.25); // Épaisseur du trait en points (1 point = 0.3528 mm)
    doc.rect(marge+2, marge+2, largeurPage-4 - (2 * marge), hauteurPage-4 - (2 * marge), 'D'); // Position x, y, largeur, hauteur

    //#endregion
    
    const lignesTitre = splitTextIntoLines(offre.titre, largeurPage - (2 * marge+20)-20, 39, 'Arial');
    let yTitre = 50;
    for (const line of lignesTitre) {
      doc.setFont('Arial', 'bold');
      doc.setFontSize(39)
      const largeurTitre = doc.getTextWidth(line);
      const xTitre = (297 - largeurTitre) / 2; // Centrer horizontalement
      doc.text(line, xTitre, yTitre);
      yTitre += 15; // Augmenter l'espacement entre les lignes
    }

    nouveauTexteCenter(doc, 'Arial', 'normal', 30, offre.lieu, 100)
    nouveauTexteCenter(doc, 'Arial', 'bold', 40, `Offre n°${offre.numero}`, 120)

    let contenuOffre = offre.type;
    if (offre.duree && offre.duree.trim() !== '') { // Vérifie si dureeOffre n'est pas null ou vide
        contenuOffre += ` - ${offre.duree}`;
    }
    if (contenuOffre) {
    doc.setFont('Arial', 'bold');
    doc.setFontSize(25);
    const largeurContenuOffre = doc.getTextWidth(contenuOffre);
    const xContenuOffre = 60; // Aligner à gauche
    doc.text(contenuOffre, xContenuOffre, 185);

    }

    var imgData = 'data:image/png;base64,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';
    doc.addImage(imgData, 'PNG', 20, 155, 35, 35);

    if (qrCodeLink + offre.numero) {
        const qrCodeX = 240; // Définissez la position X du QR code sur la page
        const qrCodeY = 155; // Définissez la position Y du QR code sur la page
        const qrCodeSize = 6 * 5.9161; // Définissez la taille du QR code
        await addQRCodeToPage(doc, qrCodeLink + offre.numero, qrCodeX, qrCodeY, qrCodeSize);
    }
    if (i !== sheetData.length - 1) {
      doc.addPage(); // Ajouter une nouvelle page pour toutes les itérations sauf la dernière
    }
  }



  const cadrePDF = document.getElementById('cadrePDF');
  cadrePDF.innerHTML = '';
  const iframe = document.createElement('iframe');
  iframe.src = doc.output('datauristring');
  iframe.style.width = '100%';
  iframe.style.height = '100%';
  iframe.style.border = 'none';

  cadrePDF.appendChild(iframe);
    
};
