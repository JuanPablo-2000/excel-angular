import { Component } from '@angular/core';
import * as ExcelJS from 'exceljs';
import * as FileSaver from 'file-saver';
import  '../../assets/univalle';

@Component({
  selector: 'app-export-excel',
  template: '<button (click)="exportExcel()">Exportar Excel</button>',
})
export class ExportExcelComponent {

  exportExcel(): void {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');

    // Ajuste de columnas
    worksheet.columns = [
      { key: 'A', width: 10 },
      { key: 'B', width: 10 },
      { key: 'C', width: 30 },
      { key: 'D', width: 20 },
      { key: 'E', width: 20 },
      { key: 'F', width: 20 },
      { key: 'G', width: 30 },
      { key: 'H', width: 20 },
    ];

    // Imagen de la universidad
    const imageId = workbook.addImage({
      base64: '../../assets/univalle',
      extension: 'jpeg',
    });
    worksheet.addImage(imageId, 'C2:C4');

    // Título del proyecto
    worksheet.mergeCells('D2:H2');
    const titleCell = worksheet.getCell('D2');
    titleCell.value = 'PROYECTO DE ASIGNACION ACADEMICA';
    titleCell.font = { name: 'Arial', size: 16, bold: true };
    titleCell.alignment = { horizontal: 'center' };

    // Tabla del periodo académico
    worksheet.mergeCells('F5:G5');
    const periodoCell = worksheet.getCell('F5');
    periodoCell.value = 'PERIODO ACADEMICO';
    periodoCell.font = { name: 'Arial', size: 8 };
    periodoCell.alignment = { horizontal: 'center' };
    periodoCell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    worksheet.mergeCells('F6:G6');
    const fechaCell = worksheet.getCell('F6');
    fechaCell.value = '2024-1';
    fechaCell.font = { name: 'Arial', size: 8 };
    fechaCell.alignment = { horizontal: 'center' };
    fechaCell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    const totalHorasCell = worksheet.getCell('H5');
    totalHorasCell.value = 'TOTAL HORAS';
    totalHorasCell.font = { name: 'Arial', size: 8 };
    totalHorasCell.alignment = { horizontal: 'center' };
    totalHorasCell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    const totalHorasValueCell = worksheet.getCell('H6');
    totalHorasValueCell.value = 1697.5;
    totalHorasValueCell.font = { name: 'Arial', size: 8 };
    totalHorasValueCell.alignment = { horizontal: 'center' };
    totalHorasValueCell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    // Identificación del docente
    worksheet.mergeCells('C7:H7');
    const docenteTitleCell = worksheet.getCell('C7');
    docenteTitleCell.value = '1. IDENTIFICACION DEL DOCENTE';
    docenteTitleCell.font = { name: 'Arial', size: 12, bold: true };
    docenteTitleCell.alignment = { horizontal: 'center' };

    // Titulos de la tabla de identificación
    const titles = ['CEDULA', 'NOMBRE', '1 APELLIDO', '2 APELLIDO', 'UNIDAD ACADEMICA'];
    const titleCells = ['C8', 'D8', 'E8', 'F8', 'G8:H8'];
    titleCells.forEach((cell, index) => {
      if (cell.includes(':')) {
        worksheet.mergeCells(cell);
      }
      const titleCell = worksheet.getCell(cell.split(':')[0]);
      titleCell.value = titles[index];
      titleCell.font = { name: 'Arial', bold: true };
      titleCell.alignment = { horizontal: 'center' };
      titleCell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    // Datos del docente
    const data = ['1116247617', 'ROYER', 'ESTRADA', 'APONTE', 'SECCIONAL TULUA'];
    const dataCells = ['C9', 'D9', 'E9', 'F9', 'G9:H9'];
    dataCells.forEach((cell, index) => {
      if (cell.includes(':')) {
        worksheet.mergeCells(cell);
      }
      const dataCell = worksheet.getCell(cell.split(':')[0]);
      dataCell.value = data[index];
      dataCell.font = { name: 'Arial' };
      dataCell.alignment = { horizontal: 'center' };
      dataCell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    // Titulos y datos adicionales
    const additionalTitles = ['VINCULACION', 'CATEGORIA', 'DEDICACION', 'NIVEL ALCANZADO', 'CENTRO COSTO'];
    const additionalTitleCells = ['C11', 'D11', 'E11', 'F11', 'G11:H11'];
    additionalTitleCells.forEach((cell, index) => {
      if (cell.includes(':')) {
        worksheet.mergeCells(cell);
      }
      const titleCell = worksheet.getCell(cell.split(':')[0]);
      titleCell.value = additionalTitles[index];
      titleCell.font = { name: 'Arial', bold: true };
      titleCell.alignment = { horizontal: 'center' };
      titleCell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    const additionalData = ['OCASIONAL', 'ASOCIADO', 'TC', 'MAESTRIA', ''];
    const additionalDataCells = ['C12', 'D12', 'E12', 'F12', 'G12:H12'];
    additionalDataCells.forEach((cell, index) => {
      if (cell.includes(':')) {
        worksheet.mergeCells(cell);
      }
      const dataCell = worksheet.getCell(cell.split(':')[0]);
      dataCell.value = additionalData[index];
      dataCell.font = { name: 'Arial' };
      dataCell.alignment = { horizontal: 'center' };
      dataCell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      FileSaver.saveAs(blob, 'AsignacionAcademica.xlsx');
    });
  }
}

