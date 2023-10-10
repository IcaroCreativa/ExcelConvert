document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  const convertButton = document.getElementById("convertButton");
  let excelFile = null;
  const fileInputTest = document.getElementById("testFile");

  const fileNameDisplay = document.getElementById("fileNameDisplay");
  const participantCountDisplay = document.getElementById("participantCount");
  const authorNameDisplay = document.getElementById("authorName");
  const creationDateDisplay = document.getElementById("creationDate");
  const lastModifiedDateDisplay = document.getElementById("lastModifiedDate");
  var fileNameToSave = "";

  // À l'intérieur de votre gestionnaire d'événements pour fileInputTest
  fileInputTest.addEventListener("change", (e) => {
    e.preventDefault();
    const wb = new ExcelJS.Workbook();
    const reader = new FileReader();

    const fileToConvert = e.target.files[0];
    if (fileToConvert) {
      const fileName = fileToConvert.name;
      fileNameToSave = `${fileName.split(".")[0]}`;
      fileNameDisplay.textContent = `${fileName.split(".")[0]}`;
      const fileExtension = fileName.split(".").pop();
      if (fileExtension === "xls" || fileExtension === "xlsx") {
        // affiche la carte avec les infos et le fichier excel
        const cardContainer = document.getElementById("card_container");
        cardContainer.classList.remove("hidden");

        const reader = new FileReader();
        reader.onload = (event) => {
          const arrayBuffer = event.target.result;
          excelFile = arrayBuffer;
        };
        reader.readAsArrayBuffer(fileToConvert);
      } else {
        console.error("Veuillez sélectionner un fichier Excel valide.");
        alert("Veuillez sélectionner un fichier Excel valide.");
      }
    }

    var file = document.getElementById("testFile");
    reader.readAsArrayBuffer(file.files[0]);

    reader.onload = () => {
      const buffer = reader.result;
      wb.xlsx.load(buffer).then((workbook) => {
        const excelDataDiv = document.getElementById("excelData");
        excelDataDiv.innerHTML = ""; // Efface tout contenu précédent
        workbook.eachSheet((sheet, id) => {
          const sheetDiv = document.createElement("div");
          sheetDiv.textContent = `Feuille ${id}:`;

          const table = document.createElement("table");
          table.classList.add(
            "w-full",
            "table-auto",
            "border-collapse",
            "border"
          );
          const tableHead = document.createElement("thead");
          const tableBody = document.createElement("tbody");

          let isHeaderRow = true; // Indicateur pour la première ligne

          // Ajoutez les données de chaque ligne de la feuille à la table
          sheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
            const tableRow = document.createElement("tr");
            row.eachCell({ includeEmpty: true }, (cell, colIndex) => {
              const cellElement = isHeaderRow
                ? document.createElement("th")
                : document.createElement("td");
              const cellValue = cell.value;

              // Formate la date au format "12/02/2012"
              if (cell.type === ExcelJS.ValueType.Date) {
                const date = new Date(cellValue);
                const day = String(date.getDate()).padStart(2, "0");
                const month = String(date.getMonth() + 1).padStart(2, "0");
                const year = date.getFullYear();
                cellElement.textContent = `${day}/${month}/${year}`;
              } else {
                cellElement.textContent = cellValue || ""; // Assurez-vous que le texte est vide s'il n'y a pas de valeur
              }

              cellElement.classList.add("border", "border-gray-500", "p-2"); // Ajoutez des classes Tailwind pour la bordure et le padding

              // Appliquez un fond violet clair à la première cellule de l'en-tête
              if (isHeaderRow) {
                cellElement.classList.add("bg-purple-100"); // Couleur violette claire
              }

              tableRow.appendChild(cellElement);
            });

            if (isHeaderRow) {
              tableHead.appendChild(tableRow); // Ajoutez la première ligne en tant qu'en-tête
              isHeaderRow = false; // Ne traitez plus les lignes comme des en-têtes
            } else {
              tableBody.appendChild(tableRow); // Ajoutez les autres lignes au corps de la table
            }
          });

          table.appendChild(tableHead);
          table.appendChild(tableBody);

          sheetDiv.appendChild(table);
          excelDataDiv.appendChild(sheetDiv);

          // Mettez à jour les informations supplémentaires
          participantCountDisplay.textContent =
            workbook.worksheets[0].rowCount - 1;
          authorNameDisplay.textContent = workbook.creator || "N/A";
          const creationDate = workbook.created;
          if (creationDate instanceof Date) {
            const day = String(creationDate.getDate()).padStart(2, "0");
            const month = String(creationDate.getMonth() + 1).padStart(2, "0");
            const year = creationDate.getFullYear();
            creationDateDisplay.textContent = `${day}/${month}/${year}`;
          } else {
            creationDateDisplay.textContent = "N/A";
          }
          lastModifiedDateDisplay.textContent =
            workbook.lastModifiedBy || "N/A";
        });
      });
    };
  });

  convertButton.addEventListener("click", async () => {
    if (!excelFile) {
      console.error("Veuillez sélectionner un fichier Excel valide.");
      alert("Veuillez sélectionner un fichier Excel valide.");
      return;
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "AppSim";
    workbook.lastModifiedBy = "AppSim";
    workbook.created = new Date();

    try {
      await workbook.xlsx.load(excelFile);

      const worksheet = workbook.getWorksheet(1);

      // Supprimez la première ligne du tableau
      worksheet.spliceRows(1, 1);

      // Convertissez la deuxième colonne de "MME" à "Mme"
      worksheet.eachRow((row) => {
        const cell = row.getCell(2);
        if (cell.value === "MME") {
          cell.value = "Mme";
        }
      });

      // Parcourez les lignes et convertissez la cinquième colonne de type Date en texte avec le format "dd/mm/yyyy"
      worksheet.eachRow((row) => {
        const dateCell = row.getCell(5);
        if (dateCell.type === ExcelJS.ValueType.Date) {
          const date = dateCell.value;
          const day = String(date.getDate()).padStart(2, "0");
          const month = String(date.getMonth() + 1).padStart(2, "0");
          const year = date.getFullYear();
          const formattedDate = `${day}/${month}/${year}`;
          row.getCell(6).value = formattedDate; // Nouvelle colonne à la position 6
        }
      });

      // Supprimez la colonne de dates d'origine (colonne 5)
      worksheet.spliceColumns(5, 1);

      // Convertissez le classeur en un fichier blob
      const buffer = await workbook.xlsx.writeBuffer();

      // Créez un Blob à partir du buffer
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      // Créez un URL pour le blob
      const url = window.URL.createObjectURL(blob);

      // Créez un lien de téléchargement
      const a = document.createElement("a");
      a.href = url;
      a.download = `${fileNameToSave}_converted.xlsx`;

      // Cliquez sur le lien pour démarrer le téléchargement
      a.click();
    } catch (error) {
      console.error("Erreur lors de la conversion du fichier Excel :", error);
      alert(
        "Erreur lors de la conversion du fichier Excel. Vérifiez que le format du fichier est .xls"
      );
    }
  });
});
