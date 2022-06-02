package br.com.luciano.apachepoi.model.exel;

import br.com.luciano.apachepoi.model.Cliente;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.IndexedColors;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class CriarArquivoExcel {

    public void criarArquivo(String nomeArquivo, List<Cliente> clientes){


        try(XSSFWorkbook workbook = new XSSFWorkbook(); FileOutputStream outputStream = new FileOutputStream(nomeArquivo)){

            XSSFSheet sheet = workbook.createSheet("Lista de equipes");

            insertIcons(workbook, "/home/luciano/Downloads/caixa.png", sheet, 0, 6, 0, 3);

            int numeroLinha = 4;
            XSSFFont fonteNegrito = this.criarFonte(workbook, true);
            XSSFFont fonteSemNegrito = this.criarFonte(workbook, false);
            CellStyle cellStyleCentro = this.criarCellStyle(workbook, HorizontalAlignment.CENTER);
            CellStyle cellStyleEsquerda = this.criarCellStyle(workbook, HorizontalAlignment.LEFT);

            this.criarCelulaComTitulo(numeroLinha++, sheet, "Extrato clubes de futebol", cellStyleCentro, "$A$5:$I5", 8,fonteNegrito);
            this.criarCelulaComTitulo(numeroLinha++, sheet, "18/02/2021 HR Brasilia 17:47:09", cellStyleCentro, "$A$6:$I$6", 8, fonteNegrito);
            this.criarCelulaComTitulo(numeroLinha++, sheet, "Razão social: Sociedade esportiva gama", cellStyleCentro, "$A$7:$I$7", 8, fonteNegrito);
            this.criarCelulaComTitulo(numeroLinha++, sheet, "Nome Fantasia: GAMA/DF -CNPJ 00.442.129/0001-50", cellStyleCentro, "$A$8:$I$8", 8, fonteNegrito);
            this.criarCelulaComTitulo(numeroLinha++, sheet, "Data de referencia do movimento: Janeiro 2021", cellStyleCentro, "$A$9:$I$9", 8, fonteNegrito);

            //Pulando uma linha
            numeroLinha++;

            this.criarCelulaComTitulo(numeroLinha++, sheet, "DEMONSTRATIVO ARRECADAÇÃO DA TIMEMANIA - CONCURSO 1584 A 1595",
                    cellStyleCentro, "$A$11:$I$11", 8, fonteNegrito);

            Row row = this.criarCelulaComTitulo(numeroLinha++, sheet, "Descrição", cellStyleCentro, "$A$12:$G$12", 6, fonteNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "Valor (R$)","$H$12:$I$12", fonteNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "Arrecadação Total", cellStyleEsquerda, "$A$13:$G$13", 6, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "17.051.787,00","$H$13:$I$13", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "Rateio para Clubes Grupos I e II (11%)", cellStyleEsquerda, "$A$14:$G$14", 6, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "1.875.696,57","$H$14:$I$14", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "Rateio para Time do Coração (11%)", cellStyleEsquerda, "$A$15:$G$15", 6, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "1.875.696,57","$H$15:$I$15", fonteSemNegrito, 7);

            //Pulando uma linha
            numeroLinha++;

            this.criarCelulaComTitulo(numeroLinha++, sheet, "QUADRO DE REPASSE - JANEIRO 2021",
                    cellStyleCentro, "$A$17:$I$17", 8, fonteNegrito);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "Produto", cellStyleCentro, "$A$18:$G$18", 6, fonteNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "Valor (R$)","$H$18:$I$18", fonteNegrito,7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "TIMEMANIA - Distribuição Grupos I e II", cellStyleEsquerda, "$A$19:$G$19", 6, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "14.428,43","$H$19:$I$19", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "TIMEMANIA - Time do Coração (0.937570", cellStyleEsquerda, "$A$20:$G$20", 6, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "17.585,96","$H$20:$I$20", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "TIMEMANIA - Redidtribuição Conforme art 4 , 3 Decreto N 10.941, de 13 de Janeiro de 2022", cellStyleEsquerda, "$A$21:$G$21", 6, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "0,00","$H$21:$I$21", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "LOTECA", cellStyleEsquerda, "$A$22:$G$22", 6, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "0,00","$H$22:$I$22", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "LOTOGOL", cellStyleEsquerda, "$A$23:$G$23", 6, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "0,00","$H$23:$I$23", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "SUBTOTAL", cellStyleCentro, "$A$24:$G$24", 6, fonteNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "32.014,39","$H$24:$I$24", fonteNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "IR (9,45%)", cellStyleEsquerda, "$A$25:$G$25", 6, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "3.025,36","$H$25:$I$25", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "INSS -(5%)", cellStyleEsquerda, "$A$26:$G$26", 6, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "1.600,92","$H$26:$I$26", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "TOTAL", cellStyleCentro, "$A$27:$G$27", 6, fonteNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "27.388,11","$H$27:$I$27", fonteNegrito, 7);

            //Pulando uma linha
            numeroLinha++;

            this.criarCelulaComTitulo(numeroLinha++, sheet, "DISTRIBUIÇÃO DO REPASSE",
                    cellStyleCentro, "$A$29:$I$29", 8, fonteNegrito);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "Beneficiário", cellStyleCentro, "$A$30:$E$30", 4, fonteNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "Percentual","$F$30:$G$30", fonteNegrito, 5);
            this.criarCelulaConValor(row, workbook, sheet, "Valor (R$)","$H$30:$I$30", fonteNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "Pagamento Mandato Judicial", cellStyleEsquerda, "$A$31:$E$31", 4, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "100%","$F$31:$G$31", fonteSemNegrito, 5);
            this.criarCelulaConValor(row, workbook, sheet, "27.388,11","$H$31:$I$31", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "BLOQUEIO -Trazer o motivo do bloqueio", cellStyleEsquerda, "$A$32:$E$32", 4, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "0,00%","$F$32:$G$32", fonteSemNegrito, 5);
            this.criarCelulaConValor(row, workbook, sheet, "0,00","$H$32:$I$32", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "FGTS - Fundo de Garantia por Tempo de serviço", cellStyleEsquerda, "$A$33:$E$33", 4, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "0,00%","$F$33:$G$33", fonteSemNegrito, 5);
            this.criarCelulaConValor(row, workbook, sheet, "0,00","$H$33:$I$33", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "PGFN - Fazenda Nacional", cellStyleEsquerda, "$A$34:$E$34", 4, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "0,00%","$F$34:$G$34", fonteSemNegrito, 5);
            this.criarCelulaConValor(row, workbook, sheet, "0,00","$H$34:$I$34", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "INSS - Receita Previdenciária", cellStyleEsquerda, "$A$35:$E$35", 4, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "0,00%","$F$35:$G$35", fonteSemNegrito, 5);
            this.criarCelulaConValor(row, workbook, sheet, "0,00","$H$35:$I$35", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "RFB - Receita Tributária", cellStyleEsquerda, "$A$36:$E$36", 4, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "0,00%","$F$36:$G$36", fonteSemNegrito, 5);
            this.criarCelulaConValor(row, workbook, sheet, "0,00","$H$36:$I$36", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "Conta Clube - Caixa", cellStyleEsquerda, "$A$37:$E$37", 4, fonteSemNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "0,00%","$F$37:$G$37", fonteSemNegrito, 5);
            this.criarCelulaConValor(row, workbook, sheet, "0,00","$H$37:$I$37", fonteSemNegrito, 7);

            row = this.criarCelulaComTitulo(numeroLinha++, sheet, "TOTAL", cellStyleCentro, "$A$38:$E$38", 4, fonteNegrito);
            this.criarCelulaConValor(row, workbook, sheet, "100%","$F$38:$G$38", fonteNegrito, 5);
            this.criarCelulaConValor(row, workbook, sheet, "27,388,11","$H$38:$I$38", fonteNegrito, 7);

            workbook.write(outputStream);

        }catch(IOException ex){
            ex.printStackTrace();
        }
    }

    private Row criarCelulaComTitulo(int numeroLinha, XSSFSheet sheet, String valor, CellStyle cellStyle, String alcance,
                                     int numeroColunas, XSSFFont fonte){

        Row row = sheet.createRow(numeroLinha);

        Cell cellTitulo = row.createCell(0);
        cellTitulo.setCellValue(valor);

        cellStyle.setFont(fonte);
        cellTitulo.setCellStyle(cellStyle);
        sheet.addMergedRegion(CellRangeAddress.valueOf(alcance));

        for(int i = 1; i <= numeroColunas; i++){
            Cell cell = row.createCell(i);
            cell.setCellStyle(cellStyle);
        }
        return row;
    }

    private void criarCelulaConValor(Row row, XSSFWorkbook workbook, XSSFSheet sheet, String valor,  String alcance, XSSFFont fonte, int colunaInicial){

        short black = IndexedColors.BLACK.getIndex();

        //Criando o estilo na linha 7
        CellStyle styleColuna7 = workbook.createCellStyle();
        styleColuna7.setBorderBottom(BorderStyle.THIN);
        styleColuna7.setBottomBorderColor(black);
        styleColuna7.setAlignment(HorizontalAlignment.RIGHT);
        styleColuna7.setFont(fonte);

        Cell cell7 = row.createCell(colunaInicial);
        cell7.setCellStyle(styleColuna7);
        cell7.setCellValue(valor);

        sheet.addMergedRegion(CellRangeAddress.valueOf(alcance));

        //Criando o estilo na linha 8
        CellStyle styleColuna8 = workbook.createCellStyle();
        styleColuna8.setBorderRight(BorderStyle.THIN);
        styleColuna8.setRightBorderColor(black);
        styleColuna8.setBorderBottom(BorderStyle.THIN);
        styleColuna8.setBottomBorderColor(black);

        Cell cell8 = row.createCell(++colunaInicial);
        cell8.setCellStyle(styleColuna8);


    }

    private CellStyle criarCellStyle(XSSFWorkbook workbook, HorizontalAlignment horizontalAlignment){
        short black = IndexedColors.BLACK.getIndex();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(horizontalAlignment);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(black);

        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(black);

        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(black);

        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(black);
        return cellStyle;
    }

    private void insertIcons(XSSFWorkbook workbook, String URL, XSSFSheet sheet, int colBegin, int colEnd, int rowBegin, int rowEnd) {
        try {
            InputStream iconInput = new FileInputStream(URL);
            byte[] byteTransf = IOUtils.toByteArray(iconInput);
            int pictureIdx = workbook.addPicture(byteTransf, Workbook.PICTURE_TYPE_PNG);
            iconInput.close();

            CreationHelper helper = workbook.getCreationHelper();
            Drawing drawingIcon = sheet.createDrawingPatriarch();

            ClientAnchor anchorIcon = helper.createClientAnchor();
            anchorIcon.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
            anchorIcon.setCol1(colBegin);
            anchorIcon.setCol2(colEnd);
            anchorIcon.setRow1(rowBegin);
            anchorIcon.setRow2(rowEnd);

            Picture iconReady = drawingIcon.createPicture(anchorIcon, pictureIdx);
            iconReady.resize(1);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private XSSFFont criarFonte(XSSFWorkbook workbook, boolean bold){
        XSSFFont fonte = workbook.createFont();
        fonte.setFontHeight(7.5);
        fonte.setFontName("Arial");
        fonte.setBold(bold);
        return fonte;
    }
}
