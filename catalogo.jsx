// Rutas configurables al inicio del archivo
var csvFilePath = "C:/Users/Windows/Desktop/Abin SAS/4. Automatizacion InDesing/Catalogo/Automatizacion de InDesing/Datos.csv";
var indesignFilePath = "C:/Users/Windows/Desktop/Abin SAS/4. Automatizacion InDesing/Catalogo/Automatizacion de InDesing/catalogo.indd";
var baseImagePath = "C:/Users/Windows/Desktop/Abin SAS/4. Automatizacion InDesing/Catalogo/V2/";

// Función para leer un archivo CSV con separador ;
function readCSV(filePath) {
    var file = new File(filePath);
    if (file.exists) {
        file.open("r");
        var content = file.read();
        file.close();
        var lines = content.split("\n");
        var data = [];
        for (var i = 1; i < lines.length; i++) { // Empezar desde 1 para omitir la cabecera
            if (lines[i] !== "") {
                data.push(lines[i].split(";")); // Cambiar la coma por punto y coma
            }
        }
        return data;
    } else {
        alert("El archivo CSV no se encontró en la ruta especificada.");
        return [];
    }
}

// Leer datos del CSV
var products = readCSV(csvFilePath);
$.writeln("Productos leídos del CSV: " + products.length);

// Abrir el documento existente
var doc = app.open(new File(indesignFilePath));

// Función para encontrar un marco de texto por su contenido exacto
function findTextFrameWithExactText(searchText) {
    var allTextFrames = doc.textFrames;
    for (var i = 0; i < allTextFrames.length; i++) {
        if (allTextFrames[i].contents === searchText) {
            return allTextFrames[i];
        }
    }
    return null;
}

// Verificar que se haya abierto el documento
if (doc) {
    $.writeln("Documento InDesign abierto correctamente.");

    // Iterar sobre los productos y rellenar los marcos de texto
    for (var i = 0; i < products.length; i++) {
        var product = products[i];
        $.writeln("Procesando producto " + (i + 1));

        // Buscar y actualizar marcos de texto según su contenido
        try {
            var codeFrame = findTextFrameWithExactText("Código");
            if (codeFrame) {
                codeFrame.contents = product[0] || "";
                $.writeln("Código actualizado: " + product[0]);
            } else {
                $.writeln("No se encontró el marco de texto con el contenido 'Código'");
            }

            var refFrame = findTextFrameWithExactText("REF");
            if (refFrame) {
                refFrame.contents = product[1] || "";
                $.writeln("REF actualizado: " + product[1]);
            } else {
                $.writeln("No se encontró el marco de texto con el contenido 'REF'");
            }

            var refPremiumFrame = findTextFrameWithExactText("PREMIUM");
            if (refPremiumFrame) {
                refPremiumFrame.contents = product[2] || "";
                $.writeln("REF. PREMIUM actualizado: " + product[2]);
            } else {
                $.writeln("No se encontró el marco de texto con el contenido 'PREMIUM'");
            }

            var refPartmoFrame = findTextFrameWithExactText("PARTMO");
            if (refPartmoFrame) {
                refPartmoFrame.contents = product[3] || "";
                $.writeln("REF. PARTMO actualizado: " + product[3]);
            } else {
                $.writeln("No se encontró el marco de texto con el contenido 'PARTMO'");
            }

            var descriptionFrame = findTextFrameWithExactText("Descripción ");
            if (descriptionFrame) {
                descriptionFrame.contents = product[4] || "";
                $.writeln("Descripción actualizada: " + product[4]);
            } else {
                $.writeln("No se encontró el marco de texto con el contenido 'Descripción'");
            }

            // Colocar la imagen si se proporciona
            if (product.length > 5 && product[5] && product[5] !== "") {
                var imagePath = baseImagePath + product[5];
                var imageFile = new File(imagePath);
                if (imageFile.exists) {
                    // Encontrar el marco de texto que debería contener la imagen
                    var imageFrame = findTextFrameWithExactText("Imagen");
                    if (imageFrame) {
                        // Crear un nuevo marco de imagen en la misma posición que el marco de texto "Imagen"
                        var imageFrameRect = imageFrame.parentPage.rectangles.add();
                        imageFrameRect.geometricBounds = imageFrame.geometricBounds; // Copiar tamaño del marco existente
                        imageFrameRect.place(imageFile);
                        imageFrameRect.fit(FitOptions.PROPORTIONALLY);
                        imageFrame.remove(); // Eliminar el marco de texto "Imagen" después de colocar la imagen
                        $.writeln("Imagen colocada: " + imagePath);
                    } else {
                        $.writeln("No se encontró el marco de texto con el contenido 'Imagen'");
                    }
                } else {
                    $.writeln("La imagen no se encontró: " + imagePath);
                }
            } else {
                $.writeln("No se proporcionó una imagen para el producto " + (i + 1));
            }
        } catch (e) {
            $.writeln("Error al procesar el producto " + (i + 1) + ": " + e.message);
        }
    }

    // Guardar el documento sin cerrarlo
    try {
        doc.save();
        $.writeln("Documento guardado correctamente.");
    } catch (e) {
        alert("Error al guardar el documento: " + e.message);
    }

} else {
    alert("No se pudo abrir el archivo InDesign.");
}
