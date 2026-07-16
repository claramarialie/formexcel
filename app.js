
var selected_templates = templates.bushing_cover

$(document).ready(function () {
    updateTable()

    $(".js-example-placeholder-simple").select2({
        placeholder: "Pilih...",
        data: selected_templates
    })

    $(".input-size").on("input", updateTable);

    // $(".js-example-placeholder-simple").
})

// document.querySelectorAll(".input-size").forEach(inp => {
//     inp.addEventListener('input', function () {
//         updateTable()
//     })
// })

function updateTable() {
        table = document.getElementById("table")
        n = document.getElementById("input-n").value
        m = document.getElementById("input-m").value
        table.innerHTML = ''
        var str = '';
        for (var i = 0; i < n; i++) {
            str += '<tr>'
            for (var j = 0; j < m; j++) {
                str += `
        <td>
            <input
                type="file"
                accept="image/*"
                capture="environment"
                class="foto form-control"
                data-row="${i}"
                data-col="${j}">
        </td>`;
            }
            str += '</tr>'
        }
        console.log(str)
        table.innerHTML = str
}

async function exportExcel() {

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Foto");

    const n = Number(document.getElementById("input-n").value);
    const m = Number(document.getElementById("input-m").value);

    for (let c = 1; c <= m; c++) {
        sheet.getColumn(c).width = (selected_templates.width-5) / 7;
    }

    for (let r = 1; r <= n; r++) {
        sheet.getRow(r).height = selected_templates.heigth * 0.75;
    }

    const inputs = document.querySelectorAll(".foto");

    for (const input of inputs) {

        if (input.files.length == 0)
            continue;

        const file = input.files[0];

        const img = await compressResizeImg(file, selected_templates.width-5, selected_templates.heigth-5);

        const imageId = workbook.addImage({
            base64: img.base64,
            extension: file.type.includes("png") ? "png" : "jpeg"
        });

        const row = Number(input.dataset.row);
        const col = Number(input.dataset.col);

        sheet.addImage(imageId, {
            tl: {
                col: col,
                row: row
            },
            ext: {
                width: img.width,
                height: img.height
            }
        });
    }

    const buffer = await workbook.xlsx.writeBuffer();

    saveAs(
        new Blob([buffer], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }),
        "Foto.xlsx"
    );
}

function fileToBase64(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.readAsDataURL(file);
    })
}


function compressResizeImg(file, maxW = 150, maxH = 120) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (event) => {
            const img = new Image();

            img.onload = () => {
                let width = img.width;
                let height = img.height;

                if (width > height) {
                    if (width > maxW) {
                        height = Math.round(height * maxW / width);
                        width = maxW;
                    }
                } else {
                    if (height > maxH) {
                        width = Math.round(width * maxH / height);
                        height = maxH;
                    }
                }

                // buat canvas
                const canvas = document.createElement("canvas");
                canvas.width = width;
                canvas.height = height;

                const ctx = canvas.getContext("2d");
                ctx.drawImage(img, 0, 0, width, height);

                resolve({
                    base64: canvas.toDataURL("image/jpeg", 0.8),
                    width,
                    height
                });
            };

            img.onerror = reject;
            img.src = event.target.result;
        };

        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}


