document.addEventListener('DOMContentLoaded', function() {
  // Accès à la caméra
  navigator.mediaDevices.getUserMedia({ video: true })
    .then(function(stream) {
      var video = document.getElementById('videoElement');
      video.srcObject = stream;
    })
    .catch(function(err) {
      console.error('Accès à la caméra refusé : ', err);
    });

  // Attendre que la vidéo soit chargée
  document.getElementById('videoElement').addEventListener('canplay', function() {
    var video = document.getElementById('videoElement');
    var canvas = document.getElementById('canvasElement');
    var context = canvas.getContext('2d');

    // Définir la taille du canvas pour correspondre à la vidéo
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;

    var detectButton = document.getElementById('detectTextButton');
    if (detectButton) {
      detectButton.addEventListener('click', function() {
        context.drawImage(video, 0, 0, canvas.width, canvas.height);

        // Utilisation de Tesseract.js pour la détection de texte
        Tesseract.recognize(
          canvas,
          'eng', // Langue (par exemple, 'eng' pour l'anglais)
          { logger: m => console.log(m) } // Options (vous pouvez ajuster selon les besoins)
        ).then(({ data: { text } }) => {
          var textResult = document.getElementById('textResult');
          if (text && text.includes("Offre n°")) {
            textResult.innerText = "Texte détecté : Offre n°";
          } else {
            textResult.innerText = "Aucun texte détecté";
          }
        });
      });
    } else {
      console.error("L'élément detectTextButton n'a pas été trouvé.");
    }
  });
});