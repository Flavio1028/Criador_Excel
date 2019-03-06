package util.excel;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class TutorialJExcel {

	/**
	 * Atributo utilizado para formatar as cedulas de titulo da tabela
	 */
	private WritableCellFormat timesBoldUnderline;
	/**
	 * Atributo
	 */
	private WritableCellFormat times;
	/**
	 * Atributo com o arquivo gerado
	 */
	private String inputArquivo;
	/**
	 * 
	 */
	public File arquivo = null;

	// Exemplo de Como criar uma planilha com JXL no Java
	public void setOutputFile(String inputArquivo) {
		this.inputArquivo = inputArquivo;
	}

	// Método responsável por fazer a escrita, a inserção dos dados na planilha
	public void insere() throws IOException, WriteException {

		// Cria um novo arquivo
		arquivo = new File(inputArquivo);
		WorkbookSettings wbSettings = new WorkbookSettings();

		wbSettings.setLocale(new Locale("pt", "BR"));

		WritableWorkbook workbook = Workbook.createWorkbook(arquivo, wbSettings);
		// Define um nome para a planilha
		workbook.createSheet("Termo de Moeda", 0);
		WritableSheet excelSheet = workbook.getSheet(0);
		criaLabel(excelSheet);

		workbook.write();
		workbook.close();
		
		byte[] fileContent = Files.readAllBytes(arquivo.toPath());
		
	}

	// Método responsável pela definição das labels
	private void criaLabel(WritableSheet sheet) throws WriteException {
		CellView cv = new CellView();
		cv.setFormat(times);
		cv.setFormat(timesBoldUnderline);
		cv.setAutosize(true);

		this.criarLabelsCabecalho(sheet);

		addLabel(sheet, 3, 5, "CLIENTE");
		addLabel(sheet, 4, 5, "BRADESCO");
		addLabel(sheet, 5, 5, "Simples");
		addLabel(sheet, 6, 5, "");

		addLabel(sheet, 3, 7, "27/04/2018");
		addLabel(sheet, 4, 7, "27/04/2018");
		addLabel(sheet, 5, 7, "27/04/2018");
		addLabel(sheet, 6, 7, "103");
		
		addLabel(sheet, 3, 9, "220 - DOLAR EUA");
		addLabel(sheet, 4, 9, "790 - REAL");
		addLabel(sheet, 5, 9, "100.000,00");
		addLabel(sheet, 6, 9, "3,46920000");
		
		addLabel(sheet, 3, 11, "");
		addLabel(sheet, 4, 11, "");
		addLabel(sheet, 5, 11, "");
		addLabel(sheet, 6, 11, "");
		
		addLabel(sheet, 3, 13, "");
		addLabel(sheet, 4, 13, "");
		addLabel(sheet, 5, 13, "");
		addLabel(sheet, 6, 13, "");
		
		addLabel(sheet, 3, 15, "");
		addLabel(sheet, 4, 15, "SISBACEN");
		addLabel(sheet, 5, 15, "");
		addLabel(sheet, 6, 15, "D-1");
		
		addLabel(sheet, 3, 17, "Fechamento");
		addLabel(sheet, 4, 17, "CETIP");
		addLabel(sheet, 5, 17, "");
		addLabel(sheet, 6, 17, "");
		
	}

	// Adiciona cabecalho
	private void addCabecalhio(WritableSheet planilha, int coluna, int linha, String s)
			throws RowsExceededException, WriteException {
		// Cria a fonte em negrito com underlines
		WritableFont times10ptBoldUnderline = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, false);
		// UnderlineStyle.SINGLE);
		timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
		// Efetua a quebra automática das células
		timesBoldUnderline.setWrap(Boolean.TRUE);
		//
		timesBoldUnderline.setAlignment(Alignment.CENTRE);
		//
		timesBoldUnderline.setBackground(Colour.GRAY_25);
		//
		timesBoldUnderline.setBorder(Border.ALL, BorderLineStyle.THIN);
		Label label;
		label = new Label(coluna, linha, s, timesBoldUnderline);
		planilha.addCell(label);
	}

	private void addLabel(WritableSheet planilha, int coluna, int linha, String s)
			throws WriteException, RowsExceededException {
		// Cria o tipo de fonte como TIMES e tamanho
		WritableFont times10pt = new WritableFont(WritableFont.ARIAL, 10);
		// Define o formato da célula
		times = new WritableCellFormat(times10pt);
		// Efetua a quebra automática das células
		times.setWrap(true);
		//
		times.setAlignment(Alignment.CENTRE);
		//
		times.setBorder(Border.ALL, BorderLineStyle.THIN);
		Label label;
		label = new Label(coluna, linha, s, times);
		planilha.addCell(label);
	}

	/**
	 * 
	 * @param sheet
	 */
	private void criarLabelsCabecalho(WritableSheet sheet) {
		try {
			addCabecalhio(sheet, 3, 4, "Comprador (Moeda Base):");
			sheet.setColumnView(3, 30);
			addCabecalhio(sheet, 4, 4, "Vendedor (Moeda Base):");
			sheet.setColumnView(4, 30);
			addCabecalhio(sheet, 5, 4, "Tipo do Termo:");
			sheet.setColumnView(5, 30);
			addCabecalhio(sheet, 6, 4, "");
			sheet.setColumnView(6, 30);
			// Segunda linha cabecalho
			addCabecalhio(sheet, 3, 6, "Data da contratação:");
			sheet.setColumnView(3, 30);
			addCabecalhio(sheet, 4, 6, "Data de Início efetivo:");
			sheet.setColumnView(4, 30);
			addCabecalhio(sheet, 5, 6, "Data de Vencimento:");
			sheet.setColumnView(5, 30);
			addCabecalhio(sheet, 6, 6, "Prazo (em dias):");
			sheet.setColumnView(6, 30);
			// Terceira linha cabecalho
			addCabecalhio(sheet, 3, 8, "Moeda Base (Referência):");
			sheet.setColumnView(3, 30);
			addCabecalhio(sheet, 4, 8, "Moeda Cotada:");
			sheet.setColumnView(4, 30);
			addCabecalhio(sheet, 5, 8, "Valor Moeda de Ref.:");
			sheet.setColumnView(5, 30);
			addCabecalhio(sheet, 6, 8, "Taxa de Câmbio a Termo:");
			sheet.setColumnView(6, 30);
			// Quarta linha cabecalho
			addCabecalhio(sheet, 3, 10, "Termo a Termo Tipo:");
			sheet.setColumnView(3, 30);
			addCabecalhio(sheet, 4, 10, "Tipo do Limite:");
			sheet.setColumnView(4, 30);
			addCabecalhio(sheet, 5, 10, "Tipo do Limite:");
			sheet.setColumnView(5, 30);
			addCabecalhio(sheet, 6, 10, "Limite Inferior:");
			sheet.setColumnView(6, 30);
			// Quinta linha cabecalho
			addCabecalhio(sheet, 3, 12, "Ajuste de Taxa Responsável:");
			sheet.setColumnView(3, 30);
			addCabecalhio(sheet, 4, 12, "Ajuste de Taxa Dt. Inicial:");
			sheet.setColumnView(4, 30);
			addCabecalhio(sheet, 5, 12, "Ajuste de Taxa Dt. Inicial:");
			sheet.setColumnView(5, 30);
			addCabecalhio(sheet, 6, 12, "Prêmio a ser pago:");
			sheet.setColumnView(6, 30);
			// Sexta linha cabecalho
			addCabecalhio(sheet, 3, 14, "Prêmio a ser pago Valor:");
			sheet.setColumnView(3, 30);
			addCabecalhio(sheet, 4, 14, "Fonte de Informação:");
			sheet.setColumnView(4, 30);
			addCabecalhio(sheet, 5, 14, "Fonte de Consulta:");
			sheet.setColumnView(5, 30);
			addCabecalhio(sheet, 6, 14, "Cotação de vencimento:");
			sheet.setColumnView(6, 30);
			// Setima linha cabecalho
			addCabecalhio(sheet, 3, 16, "Horário do Boletim:");
			sheet.setColumnView(3, 30);
			addCabecalhio(sheet, 4, 16, "Local de Registro:");
			sheet.setColumnView(4, 30);
			addCabecalhio(sheet, 5, 16, "");
			sheet.setColumnView(5, 30);
			addCabecalhio(sheet, 6, 16, "");
			sheet.setColumnView(6, 30);

		} catch (Exception e) {
			e.getMessage();
		}
	}
}