
import { read, utils, writeFile } from "xlsx";

const file = Bun.file("DWH-CS-SGC00050253.xlsx")
const workbook = read(Buffer.from(await file.arrayBuffer()), { type: 'buffer' });
const fold_1 = workbook.Sheets[workbook.SheetNames[0]];
const data: any[] = utils.sheet_to_json(fold_1)
const camp_scopes = ['hs_tsfecha', 'hs_hora', 'hs_oficina', 'hs_tipo_chequera', 'hs_usuario', 'hs_monto', 'hs_nombre', 'hs_referencia', 'hs_campo_alt_uno', 'hs_campo_alt_dos', 'monto_transaccion', 'hs_valor', 'hs_fonres_iess'];
const data_fill = []

for (const row of data) {
    const data_filtered: any = {}
    for (const scope of camp_scopes) {
        if (scope in row) data_filtered[scope] = row[scope]
    }
    data_fill.push(data_filtered);
}

const create_h = ['FECHA', 'HORA', 'AGENCIA', 'CIUDAD', 'CANAL', 'COD CAJERO', 'MONTO TOTAL', 'EFECTIVO', 'CHEQUE', 'N/D', 'T/C', 'NOMBRE AFILIADO', 'CEDULA-RUC', 'CODIGO DEUDA', 'VALOR DEUDA', 'VALOR A PAGAR', 'COMPROBANTE']

const get_office = (str: string) => {
    switch (str) {
        case '0':
            return "MATRIZ"
        case '93':
            return "LOS CEIBOS"
        default:
            break;
    }
}

const get_city = (str: string) => {
    switch (str) {
        case '0':
            return "GUAYAQUIL"
        case '93':
            return "OLMEDO"
        default:
            break;
    }
}

const get_canal = (str: string) => {
    switch (str) {
        case 'ATM':
            return 'VEINTI4EFECTIVO'
        case 'DBA':
            return 'AUTOMATICO'
        case 'IBK':
            return '24online'
        case 'IVR':
            return 'VEINTI4FONO'
        case 'KSK':
            return 'PUNTOVEINTI4'
        case 'VEN':
            return 'VENTANILLA'
        case 'WAP':
            return 'WAP'
        case 'CNB':
            return 'CNB'
        case 'SAT':
            return 'SAT'
        case 'DIR':
            return 'AUTOMATICO'
        default:
            return '';
    }
}

const data_remaped = data_fill.map((x) => ({
    'FECHA': x.hs_tsfecha,
    'HORA': x.hs_hora.split(' ')[1],
    'AGENCIA': get_office(x.hs_oficina),
    'CIUDAD': get_city(x.hs_oficina),
    'CANAL': get_canal(x.hs_tipo_chequera),
    'COD CAJERO': x.hs_usuario,
    'MONTO TOTAL': x.monto_transaccion,
    'EFECTIVO': "0.00",
    'CHEQUE': "0.00", 
    'N/D': x.hs_monto,
    'T/C': "0.00",
    'NOMBRE AFILIADO': x.hs_nombre,
    'CEDULA-RUC': x.hs_referencia,
    'CODIGO DEUDA': x.hs_campo_alt_uno,
    'VALOR DEUDA': x.hs_monto,
    'VALOR A PAGAR': x.hs_monto,
    'COMPROBANTE': x.hs_campo_alt_dos
}))



// Crear un nuevo libro de trabajo
const new_workbook = utils.book_new();

// Convertir los datos a una hoja de cálculo
const worksheet = utils.json_to_sheet(data_remaped, { header: create_h });

// Añadir la hoja al libro
utils.book_append_sheet(new_workbook, worksheet, "Hoja1");

// Escribir el archivo
writeFile(new_workbook, 'DWH-CS-SGC00050253_OUT.xlsx');




// const create_table = `${camp_scopes.join('\t').toString()}`
// const create_isert = `INSERT INTO cob_cuentas..his_temp \n (${camp_scopes.join(',').toString()}) `
// const create_values = `VALUES \n ${data_fill.map((val) => `(${Object.values(val).map((v) => `'${v}'`).join(',')}) \n`).toString()}`
// const query = `${create_table} \n ${create_isert} \n ${create_values}`

// // Bun.write("out.text", query)






console.log("Hello via Bun!\n \r", data_remaped)