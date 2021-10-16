package coconut;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
public class ProjectCoconut {


	static Scanner sc = new Scanner(System.in);
	static HSSFWorkbook workBook = new HSSFWorkbook();
	static HSSFSheet sheet = workBook.createSheet("Overview");

	public static void main(String[] args) throws Exception {
		HSSFFont font = workBook.createFont();
		HSSFCellStyle style = workBook.createCellStyle();
		heading(font, style);
		for (int i = 1; i <= 2; i++) {
			getInput(i);
		}
		try {
			FileOutputStream file = new FileOutputStream("D:\\Excel\\ProjectCoconut.xls");
			workBook.write(file);
		} catch (FileNotFoundException e) {
			System.out.println(e);
		}
	}

	private static void getInput(int i) throws Exception {
		System.out.println("Enter Month:");
		String month = sc.next();
		System.out.println("Enter no of Pices:");
		int totalPieces = sc.nextInt();
		category(i, month, totalPieces);
	}

	private static void category(int i, String month, int totalPieces) {
		System.out.println("Enter Category:");
		String category = sc.next();
		if (category.equals("ton")) {
			System.out.println("Enter Price Per Ton:");
			int pricePerTon = sc.nextInt();
			double pieceRate = piece(pricePerTon);
			positioning(i, month, totalPieces, category, pricePerTon, pieceRate);

		} else if (category.equals("piece")) {
			System.out.println("Enter Price Per Piece:");
			double pricePerPiece = sc.nextDouble();
			int tonRate = ton(pricePerPiece);
			positioning(i, month, totalPieces, category, tonRate, pricePerPiece);
		} else {
			System.out.println("Enter Correct Value!");
			category(i, month, totalPieces);
		}
	}

	private static double piece(int pricePerTon) {
		return ((double) pricePerTon / 1000);
	}

	private static int ton(double pricePerPiece) {
		return (int) (pricePerPiece * 1000);
	}

	private static void positioning(int i, String month, int totalPieces, String category, int tonRate, double PieceRate) {

		String FB = "F" + (i + 1) + "-" + "B" + (i + 1);
		String FG = "F" + (i + 1) + "-" + "G" + (i + 1);
		Row row = sheet.createRow(i);

		Cell cell = row.createCell(0);
		cell.setCellValue(month);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue(totalPieces);
		Cell cell2 = row.createCell(2);
		cell2.setCellValue(category);
		Cell cell3 = row.createCell(3);
		cell3.setCellValue(tonRate);
		Cell cell4 = row.createCell(4);
		cell4.setCellValue(PieceRate);
		Cell cell5 = row.createCell(5);
		cell5.setCellValue(totalPieces * PieceRate);
		Cell cell6 = row.createCell(6);
		cell6.setCellFormula(FB);
		Cell cell7 = row.createCell(7);
		cell7.setCellFormula(FG);
		HSSFFont font = workBook.createFont();
		HSSFCellStyle style = workBook.createCellStyle();
		if ((totalPieces * PieceRate) > 0) {
			font.setColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
			style.setFont(font);
			Cell cell8 = row.createCell(8);
			cell8.setCellValue("Profit");
			cell8.setCellStyle(style);
		} else if ((totalPieces * PieceRate) <= 0) {
			font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
			style.setFont(font);
			Cell cell8 = row.createCell(8);
			cell8.setCellValue("Loss");
			cell8.setCellStyle(style);
		}
	}

	private static void heading(HSSFFont font, HSSFCellStyle style) {
		String tittle[] = { "Month", "Total no of Pieces", "Category", "Price Per Ton", "Price Per Piece", "Total Profit",
				"Invested Amount", "Outcome" };
		font.setBold(true);
		font.setFontHeightInPoints((short) 11);
		font.setFontName("Calibri");
		style.setFont(font);

		int count = 0;
		Row row = sheet.createRow(0);
		for (String data : tittle) {
			Cell cell = row.createCell(count);
			cell.setCellValue(data);
			cell.setCellStyle(style);
			count++;
		}
	}


}
