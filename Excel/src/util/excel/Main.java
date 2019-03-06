package util.excel;

import java.io.IOException;

import jxl.write.WriteException;


public class Main {

	public static void main(String[] args) {
		TutorialJExcel exemplo = new TutorialJExcel();
		// Define o caminho e nome do arquivo que será criado
		exemplo.setOutputFile("ExemploJExcel.xls");
		try {
			exemplo.insere();
		} catch (WriteException e) {

			e.printStackTrace();
		} catch (IOException e) {

			e.printStackTrace();
		}
	}

}
