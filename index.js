const xlsx = require('xlsx');
const {
    inputFileName,
    outputFileName,
    sheetName,
    soLoNhom,
    toaDoInputSoLoThanhPham,
    toaDoInputLoNhom,
    toaDoInputTraVe,
    toaDoOutput,
    soChuSoThapPhanCuaKetQua,
} = require('./config')

const workbook = xlsx.readFile(`./${inputFileName}`);
const content = workbook.Sheets[sheetName];

const mangLoNhom = [];
const mangTraVe = [];
const soToaDoLoNhom = Number(toaDoInputLoNhom.substr(1));
const soToaDoTraVe = Number(toaDoInputTraVe.substr(1));
for (let i = 0; i < soLoNhom; i++) {
    mangLoNhom.push(Number(content[`${toaDoInputLoNhom[0]}${soToaDoLoNhom + i}`].w || 0))
    mangTraVe.push(Number(content[`${toaDoInputTraVe[0]}${soToaDoTraVe + i}`].w || 0))
}
const soLoThanhPham = Number(content[toaDoInputSoLoThanhPham].w || 0)

const tinhTongCot = (mang) => {
    return mang.reduce((prev, curr) => {
        return prev += curr
    }, 0)
}
const tongSoLuongNhom = tinhTongCot(mangLoNhom);
const tongSoTraVe = tinhTongCot(mangTraVe);
const soLuongSuDungTrungBinh = (tongSoLuongNhom - tongSoTraVe) / soLoThanhPham // eg, 5.6

const mangSoNhomDeSuDungTheoLo = []; // eg., [10, 10, 8]
for (let i = 0; i < soLoNhom; i++) {
    mangSoNhomDeSuDungTheoLo.push(mangLoNhom[i] - mangTraVe[i])
}
const mangSoNhomConLaiTheoLo = new Array(soLoThanhPham).fill(soLuongSuDungTrungBinh); // mutable, eg., [5.6, 5.6, 5.6, 5.6, 5.6]

const result = [];
for (let i = 0; i < soLoThanhPham; i++) {
    let emptyArray = [];
    for (let j = 0; j < soLoNhom; j++) {
        emptyArray.push(0)
    }
    result.push(emptyArray);
}

const round = number => number.toFixed(soChuSoThapPhanCuaKetQua);

for (let i = 0; i < soLoNhom; i++) {
    let soNhomConLai = mangSoNhomDeSuDungTheoLo[i];
    for (let j = 0; j < soLoThanhPham; j++) {
        const toBeUsed = soNhomConLai >= soLuongSuDungTrungBinh ? soLuongSuDungTrungBinh : soNhomConLai;

        if (soNhomConLai === 0) { continue }
        if (mangSoNhomConLaiTheoLo[j] >= toBeUsed) {
            result[j][i] = round(toBeUsed);
            soNhomConLai -= round(toBeUsed);
            mangSoNhomConLaiTheoLo[j] -= round(toBeUsed);
        } else {
            result[j][i] = round(mangSoNhomConLaiTheoLo[j]);
            soNhomConLai -= round(mangSoNhomConLaiTheoLo[j]);
            mangSoNhomConLaiTheoLo[j] = 0;
        }
    }
}

console.log('result:')
for (let i = 0; i < result.length; i++) {
    for (let j = 0; j < result[i].length; j++) {
        if (!Number(result[i][j])) {
            result[i][j] = undefined; // xlsx won't write in cell if value is undefined
        }
    }
    console.log(result[i])
}
// eg.,
// [ '5.60', undefined, undefined ]
// [ '4.40', '1.20', undefined ]
// [ undefined, '5.60', undefined ]
// [ undefined, '3.20', '2.40' ]
// [ undefined, undefined, '5.60' ]

xlsx.utils.sheet_add_aoa(content, result, { origin: toaDoOutput })
xlsx.writeFile(workbook, outputFileName)
console.log(`Program ran successfully. Output is in the file ${outputFileName}`)
