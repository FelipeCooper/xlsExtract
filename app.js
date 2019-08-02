var fs = require('fs');
var path = require('path')
    // Are we running app locally via node?
const isLocal = typeof process.pkg === 'undefined'

// Build the base path based on current running mode (if packaged, we need the location of executable)
const pastaAtual = isLocal ? process.cwd() : path.dirname(process.execPath) //PKG PASTA ATUAL

function getFiles(dir, files_) {
    files_ = files_ || [];
    var files = fs.readdirSync(dir);
    for (var i in files) {
        var name = dir + '/' + files[i];
        if (fs.statSync(name).isDirectory()) {
            getFiles(name, files_);
        } else {
            if (path.extname(name) == '.xlsx' && path.basename(name).substr(0, 13) == 'Ficha_tecnica') { //Busca todos arquivos 'ficha_tecnica' em xlsx
                files_.push(path.basename(name));
            }
        }
    }
    return files_;
}
(async e => {
    var nameFiles = getFiles(pastaAtual)
    var Excel = require('exceljs');
    var fichatecnica = new Excel.Workbook();
    var relatorio = new Excel.Workbook();
    let dados = []
    relatorioSheet = relatorio.addWorksheet('teste');
    for (j = 0; j < nameFiles.length; j++) {
        await fichatecnica.xlsx.readFile(nameFiles[j])
        var fichaSheet = fichatecnica.getWorksheet("FICHA");
        dados.push(Array(
            fichaSheet.getCell('C5').value,
            fichaSheet.getCell('C6').value,
            fichaSheet.getCell('E3').value,
            fichaSheet.getCell('G4').value,
            fichaSheet.getCell('G5').value
        ))
    }
    relatorioSheet.columns = [{ header: 'Sindico', key: 'sindico', width: 30, },
        { header: 'Tipo', key: 'tipo', width: 18 },
        { header: 'Cód', key: 'cod', width: 10 },

    ];
    relatorioSheet.addRows(dados)
    relatorio.xlsx.writeFile('relatorio.xlsx');
})()