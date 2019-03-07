package util.excel;

import java.io.IOException;
import jxl.write.WriteException;

public class Main {

	public static TutorialJExcel exemplo;

	public static void main(String[] args) {
		// Cria a classe que gera o arquivo
		exemplo = new TutorialJExcel();
		// Define o caminho e nome do arquivo que será criado
		exemplo.setOutputFile("arquivo_teste.xls");
		try {
			System.out.println("Iniciando criação do arquivo teste.....");
			exemplo.insere();			
			System.out.println("Arquivo gerado com sucesso.\n");
			System.out.println("O arquivo foi disponibilizado neste diretorio: " + exemplo.arquivo.getAbsolutePath());
		} catch (WriteException e) {
			System.err.println("Erro ao criar o arquivo: " + e.getMessage());
			System.exit(0);
		} catch (IOException e) {
			System.err.println("Erro ao criar o arquivo: " + e.getMessage());
			System.exit(0);
		}
	}
}