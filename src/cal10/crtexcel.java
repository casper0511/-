package cal10;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import jxl.Workbook;
import jxl.format.PageOrientation;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * �ӷ��� 
 * addNumber100() 100���ڽ�λ�ӷ� ������λ�� ʮλ1-5 
 * addNumber100_2() 100���ڲ���λ�ӷ�
 * addNumber100_3() 100���ڽ�λ�ӷ� һ����λ��+һ����λ�� ��λ
 * 
 * addNumber()  10���ڴ�10�ӷ�
 * 
 * ������ 
 * minusNumber100() 100������λ���� ������λ�� 
 * minusNumber100_2() 100���ڲ���λ����
 * minusNumber100_3() 100������λ���� һ����λ��-һ����λ�� ��λ
 * 
 * 
 **/

public class crtexcel {

	// 100���ڽ�λ�ӷ� һ����λ��
	public static List<String> addNumber100_3() {

		List<String> out = new ArrayList<String>();
		Random random = new Random();

		int ge_max = 10;
		int ge_min = 5;

		int shi_max = 5;
		int shi_min = 1;

		int a = random.nextInt(ge_max) % (ge_max - ge_min + 1) + ge_min;
		int b = random.nextInt(ge_max) % (ge_max - ge_min + 1) + ge_min;

		int shiwei = random.nextInt(shi_max) % (shi_max - shi_min + 1) + shi_min;

		out.add(String.valueOf(shiwei * 10 + a));
		out.add(String.valueOf(b));

		return out;

	}

	// 100������λ���� һ����λ��
	public static List<String> minusNumber100_3() {

		List<String> out = new ArrayList<String>();
		Random random = new Random();

		int a_max = 9;
		int a_min = 5;

		int b_max = 5;
		int b_min = 1;

		int a = random.nextInt(a_max) % (a_max - a_min + 1) + a_min;
		int b = random.nextInt(b_max) % (b_max - b_min + 1) + b_min;

		out.add(String.valueOf(a * 10 + b));
		out.add(String.valueOf(a));

		return out;

	}

	// 100���ڲ���λ����
	public static List<String> minusNumber100_2() {

		List<String> out = new ArrayList<String>();
		Random random = new Random();

		int a_max = 10;
		int a_min = 5;

		int b_max = 4;
		int b_min = 0;

		int a = random.nextInt(a_max) % (a_max - a_min + 1) + a_min;
		int b = random.nextInt(b_max) % (b_max - b_min + 1) + b_min;
		// int a = random.nextInt(10) + 1;
		// int c = a + b;

		out.add(String.valueOf(a * 10 + b));
		out.add(String.valueOf(b * 10 + a));

		return out;

	}

	// 100���ڲ���λ�ӷ�
	public static List<String> addNumber100_2() {

		List<String> out = new ArrayList<String>();
		Random random = new Random();

		// int ge_max = 10;
		// int ge_min = 5;

		int shi_max = 5;
		int shi_min = 1;

		int a = random.nextInt(shi_max) % (shi_max - shi_min + 1) + shi_min;
		int b = random.nextInt(shi_max) % (shi_max - shi_min + 1) + shi_min;

		int shiwei = random.nextInt(shi_max) % (shi_max - shi_min + 1) + shi_min;

		out.add(String.valueOf(shiwei * 10 + a));
		out.add(String.valueOf(shiwei * 10 + b));

		return out;

	}

	// 100���ڽ�λ�ӷ� ������λ�� ʮλ1-5
	public static List<String> addNumber100() {

		List<String> out = new ArrayList<String>();
		Random random = new Random();

		int ge_max = 10;
		int ge_min = 5;

		int shi_max = 5;
		int shi_min = 1;

		int a = random.nextInt(ge_max) % (ge_max - ge_min + 1) + ge_min;
		int b = random.nextInt(ge_max) % (ge_max - ge_min + 1) + ge_min;

		int shiwei = random.nextInt(shi_max) % (shi_max - shi_min + 1) + shi_min;

		out.add(String.valueOf(shiwei * 10 + a));
		out.add(String.valueOf(shiwei * 10 + b));

		return out;

	}

	// 100������λ���� ������λ��
	public static List<String> minusNumber100() {

		List<String> out = new ArrayList<String>();
		Random random = new Random();

		int a_max = 9;
		int a_min = 5;

		int b_max = 4;
		int b_min = 0;

		int a = random.nextInt(a_max) % (a_max - a_min + 1) + a_min;
		int b = random.nextInt(b_max) % (b_max - b_min + 1) + b_min;
		// int a = random.nextInt(10) + 1;
		// int c = a + b;

		out.add(String.valueOf(a * 10 + b));
		out.add(String.valueOf(b * 10 + a));

		return out;

	}

	// 10���ڴ�10�ӷ�
	public static List<String> addNumber() {

		List<String> out = new ArrayList<String>();
		Random random = new Random();

		int max = 10;
		int min = 5;

		int a = random.nextInt(max) % (max - min + 1) + min;
		int b = random.nextInt(max) % (max - min + 1) + min;

		out.add(String.valueOf(a));
		out.add(String.valueOf(b));

		return out;

	}

	// ���� ��10
	public static List<String> minusNumber() {

		List<String> out = new ArrayList<String>();
		Random random = new Random();

		int max = 10;
		int min = 5;

		int a = random.nextInt(max) % (max - min + 1) + min;
		int b = random.nextInt(10) + 1;
		int c = a + b;

		out.add(String.valueOf(c));
		out.add(String.valueOf(a));

		return out;

	}

	// 10���ڼ򵥼���
	public static List<String> minusNumber2() {

		List<String> out = new ArrayList<String>();
		Random random = new Random();

		int a = random.nextInt(10) + 1;
		int b = random.nextInt(10) + 1;
		int c = a + b;

		out.add(String.valueOf(c));
		out.add(String.valueOf(a));

		return out;

	}

	public static void createExcel(OutputStream os, int sheetnum) throws WriteException, IOException {
		// ����������
		WritableWorkbook workbook = Workbook.createWorkbook(os);

		for (int k = 0; k < sheetnum; k++) {

			String sheetname = "sheet" + k;
			// �����µ�һҳ
			WritableSheet sheet = workbook.createSheet(sheetname, 0);
			sheet.setPageSetup(PageOrientation.LANDSCAPE);
			List<String> getNumber = new ArrayList<String>();

			// �����ͷ
			sheet.mergeCells(0, 0, 29, 0);// ��Ӻϲ���Ԫ�񣬵�һ����������ʼ�У��ڶ�����������ʼ�У���������������ֹ�У����ĸ���������ֹ��
			WritableFont bold = new WritableFont(WritableFont.ARIAL, 16, WritableFont.BOLD);// ������������ͺ�����ʾ,����ΪArial,�ֺŴ�СΪ10,���ú�����ʾ
			WritableCellFormat titleFormate = new WritableCellFormat(bold);// ����һ����Ԫ����ʽ���ƶ���
			titleFormate.setAlignment(jxl.format.Alignment.CENTRE);// ��Ԫ���е�����ˮƽ�������
			titleFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);// ��Ԫ������ݴ�ֱ�������
			Label title = new Label(0, 0, "ADD & MINUS in 10", titleFormate);
			sheet.setRowView(0, 1200, false);// ���õ�һ�еĸ߶�
			sheet.addCell(title);

			WritableFont cellbold = new WritableFont(WritableFont.ARIAL, 16, WritableFont.NO_BOLD);// ������������ͺ�����ʾ,����ΪArial,�ֺŴ�СΪ10,���ú�����ʾ
			WritableCellFormat cellFormate = new WritableCellFormat(cellbold);// ����һ����Ԫ����ʽ���ƶ���
			cellFormate.setAlignment(jxl.format.Alignment.CENTRE);// ��Ԫ���е�����ˮƽ�������
			cellFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);// ��Ԫ������ݴ�ֱ�������

			// �����п�
			sheet.setColumnView(0, 4);
			sheet.setColumnView(1, 1);
			sheet.setColumnView(2, 4);
			sheet.setColumnView(3, 1);
			sheet.setColumnView(4, 10);

			sheet.setColumnView(5, 4);
			sheet.setColumnView(6, 1);
			sheet.setColumnView(7, 4);
			sheet.setColumnView(8, 1);
			sheet.setColumnView(9, 10);

			sheet.setColumnView(10, 4);
			sheet.setColumnView(11, 2);
			sheet.setColumnView(12, 4);
			sheet.setColumnView(13, 1);
			sheet.setColumnView(14, 10);

			sheet.setColumnView(15, 4);
			sheet.setColumnView(16, 1);
			sheet.setColumnView(17, 4);
			sheet.setColumnView(18, 1);
			sheet.setColumnView(19, 10);

			sheet.setColumnView(20, 4);
			sheet.setColumnView(21, 1);
			sheet.setColumnView(22, 4);
			sheet.setColumnView(23, 1);
			sheet.setColumnView(24, 10);

			sheet.setColumnView(25, 4);
			sheet.setColumnView(26, 2);
			sheet.setColumnView(27, 4);
			sheet.setColumnView(28, 1);
			sheet.setColumnView(29, 10);

			// �и�
			sheet.setRowView(0, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(1, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(2, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(3, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(4, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(5, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(6, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(7, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(8, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(9, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(10, 800, false);// ���õ�һ�еĸ߶�

			// System.out.println(sheet.getColumns());
			// System.out.println(sheet.getRows());

			// ����10��
			for (int i = 1; i < sheet.getRows(); i++) {

				getNumber = minusNumber2();

				Label label0 = new Label(0, i, getNumber.get(0), cellFormate);
				sheet.addCell(label0);
				Label label1 = new Label(1, i, "-", cellFormate);
				sheet.addCell(label1);
				Label label2 = new Label(2, i, getNumber.get(1), cellFormate);
				sheet.addCell(label2);
				Label label3 = new Label(3, i, "=", cellFormate);
				sheet.addCell(label3);

				getNumber = minusNumber();

				Label label5 = new Label(5, i, getNumber.get(0), cellFormate);
				sheet.addCell(label5);
				Label label6 = new Label(6, i, "-", cellFormate);
				sheet.addCell(label6);
				Label label7 = new Label(7, i, getNumber.get(1), cellFormate);
				sheet.addCell(label7);
				Label label8 = new Label(8, i, "=", cellFormate);
				sheet.addCell(label8);

				getNumber = addNumber();

				Label label10 = new Label(10, i, getNumber.get(0), cellFormate);
				sheet.addCell(label10);
				Label label11 = new Label(11, i, "+", cellFormate);
				sheet.addCell(label11);
				Label label12 = new Label(12, i, getNumber.get(1), cellFormate);
				sheet.addCell(label12);
				Label label13 = new Label(13, i, "=", cellFormate);
				sheet.addCell(label13);

				getNumber = minusNumber2();

				Label label15 = new Label(15, i, getNumber.get(0), cellFormate);
				sheet.addCell(label15);
				Label label16 = new Label(16, i, "-", cellFormate);
				sheet.addCell(label16);
				Label label17 = new Label(17, i, getNumber.get(1), cellFormate);
				sheet.addCell(label17);
				Label label18 = new Label(18, i, "=", cellFormate);
				sheet.addCell(label18);

				getNumber = minusNumber();

				Label label20 = new Label(20, i, getNumber.get(0), cellFormate);
				sheet.addCell(label20);
				Label label21 = new Label(21, i, "-", cellFormate);
				sheet.addCell(label21);
				Label label22 = new Label(22, i, getNumber.get(1), cellFormate);
				sheet.addCell(label22);
				Label label23 = new Label(23, i, "=", cellFormate);
				sheet.addCell(label23);

				getNumber = addNumber();

				Label label25 = new Label(25, i, getNumber.get(0), cellFormate);
				sheet.addCell(label25);
				Label label26 = new Label(26, i, "+", cellFormate);
				sheet.addCell(label26);
				Label label27 = new Label(27, i, getNumber.get(1), cellFormate);
				sheet.addCell(label27);
				Label label28 = new Label(28, i, "=", cellFormate);
				sheet.addCell(label28);
			}

		}

		// �Ѵ���������д�뵽������У����ر������
		workbook.write();
		workbook.close();
		os.close();

	}

	
	public static void createExcel100(OutputStream os, int sheetnum) throws WriteException, IOException {
		// ����������
		WritableWorkbook workbook = Workbook.createWorkbook(os);

		for (int k = 0; k < sheetnum; k++) {

			String sheetname = "sheet" + k;
			// �����µ�һҳ
			WritableSheet sheet = workbook.createSheet(sheetname, 0);
			sheet.setPageSetup(PageOrientation.LANDSCAPE);
			List<String> getNumber = new ArrayList<String>();

			// �����ͷ
			sheet.mergeCells(0, 0, 29, 0);// ��Ӻϲ���Ԫ�񣬵�һ����������ʼ�У��ڶ�����������ʼ�У���������������ֹ�У����ĸ���������ֹ��
			WritableFont bold = new WritableFont(WritableFont.ARIAL, 16, WritableFont.BOLD);// ������������ͺ�����ʾ,����ΪArial,�ֺŴ�СΪ10,���ú�����ʾ
			WritableCellFormat titleFormate = new WritableCellFormat(bold);// ����һ����Ԫ����ʽ���ƶ���
			titleFormate.setAlignment(jxl.format.Alignment.CENTRE);// ��Ԫ���е�����ˮƽ�������
			titleFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);// ��Ԫ������ݴ�ֱ�������
			Label title = new Label(0, 0, "ADD & MINUS in 100", titleFormate);
			sheet.setRowView(0, 1200, false);// ���õ�һ�еĸ߶�
			sheet.addCell(title);

			WritableFont cellbold = new WritableFont(WritableFont.ARIAL, 16, WritableFont.NO_BOLD);// ������������ͺ�����ʾ,����ΪArial,�ֺŴ�СΪ10,���ú�����ʾ
			WritableCellFormat cellFormate = new WritableCellFormat(cellbold);// ����һ����Ԫ����ʽ���ƶ���
			cellFormate.setAlignment(jxl.format.Alignment.CENTRE);// ��Ԫ���е�����ˮƽ�������
			cellFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);// ��Ԫ������ݴ�ֱ�������

			// �����п�
			sheet.setColumnView(0, 4);
			sheet.setColumnView(1, 1);
			sheet.setColumnView(2, 4);
			sheet.setColumnView(3, 1);
			sheet.setColumnView(4, 10);

			sheet.setColumnView(5, 4);
			sheet.setColumnView(6, 1);
			sheet.setColumnView(7, 4);
			sheet.setColumnView(8, 1);
			sheet.setColumnView(9, 10);

			sheet.setColumnView(10, 4);
			sheet.setColumnView(11, 2);
			sheet.setColumnView(12, 4);
			sheet.setColumnView(13, 1);
			sheet.setColumnView(14, 10);

			sheet.setColumnView(15, 4);
			sheet.setColumnView(16, 1);
			sheet.setColumnView(17, 4);
			sheet.setColumnView(18, 1);
			sheet.setColumnView(19, 10);

			sheet.setColumnView(20, 4);
			sheet.setColumnView(21, 1);
			sheet.setColumnView(22, 4);
			sheet.setColumnView(23, 1);
			sheet.setColumnView(24, 10);

			sheet.setColumnView(25, 4);
			sheet.setColumnView(26, 2);
			sheet.setColumnView(27, 4);
			sheet.setColumnView(28, 1);
			sheet.setColumnView(29, 10);

			// �и�
			sheet.setRowView(0, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(1, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(2, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(3, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(4, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(5, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(6, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(7, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(8, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(9, 800, false);// ���õ�һ�еĸ߶�
			sheet.setRowView(10, 800, false);// ���õ�һ�еĸ߶�

			// System.out.println(sheet.getColumns());
			// System.out.println(sheet.getRows());

			// ����10��
			for (int i = 1; i < sheet.getRows(); i++) {

				getNumber = addNumber100();

				Label label0 = new Label(0, i, getNumber.get(0), cellFormate);
				sheet.addCell(label0);
				Label label1 = new Label(1, i, "+", cellFormate);
				sheet.addCell(label1);
				Label label2 = new Label(2, i, getNumber.get(1), cellFormate);
				sheet.addCell(label2);
				Label label3 = new Label(3, i, "=", cellFormate);
				sheet.addCell(label3);

				getNumber = minusNumber100();

				Label label5 = new Label(5, i, getNumber.get(0), cellFormate);
				sheet.addCell(label5);
				Label label6 = new Label(6, i, "-", cellFormate);
				sheet.addCell(label6);
				Label label7 = new Label(7, i, getNumber.get(1), cellFormate);
				sheet.addCell(label7);
				Label label8 = new Label(8, i, "=", cellFormate);
				sheet.addCell(label8);

				getNumber = addNumber100_2();

				Label label10 = new Label(10, i, getNumber.get(0), cellFormate);
				sheet.addCell(label10);
				Label label11 = new Label(11, i, "+", cellFormate);
				sheet.addCell(label11);
				Label label12 = new Label(12, i, getNumber.get(1), cellFormate);
				sheet.addCell(label12);
				Label label13 = new Label(13, i, "=", cellFormate);
				sheet.addCell(label13);

				getNumber = minusNumber100_2();

				Label label15 = new Label(15, i, getNumber.get(0), cellFormate);
				sheet.addCell(label15);
				Label label16 = new Label(16, i, "-", cellFormate);
				sheet.addCell(label16);
				Label label17 = new Label(17, i, getNumber.get(1), cellFormate);
				sheet.addCell(label17);
				Label label18 = new Label(18, i, "=", cellFormate);
				sheet.addCell(label18);

				getNumber = addNumber100_3();

				Label label20 = new Label(20, i, getNumber.get(0), cellFormate);
				sheet.addCell(label20);
				Label label21 = new Label(21, i, "+", cellFormate);
				sheet.addCell(label21);
				Label label22 = new Label(22, i, getNumber.get(1), cellFormate);
				sheet.addCell(label22);
				Label label23 = new Label(23, i, "=", cellFormate);
				sheet.addCell(label23);

				getNumber = minusNumber100_3();

				Label label25 = new Label(25, i, getNumber.get(0), cellFormate);
				sheet.addCell(label25);
				Label label26 = new Label(26, i, "-", cellFormate);
				sheet.addCell(label26);
				Label label27 = new Label(27, i, getNumber.get(1), cellFormate);
				sheet.addCell(label27);
				Label label28 = new Label(28, i, "=", cellFormate);
				sheet.addCell(label28);
			}

		}

		// �Ѵ���������д�뵽������У����ر������
		workbook.write();
		workbook.close();
		os.close();

	}
	
	public static void main(String[] args) {

		try {
			OutputStream out = new FileOutputStream("C:/test1.xls");
			createExcel(out, 1);
			
			OutputStream out100 = new FileOutputStream("C:/test100.xls");
			createExcel100(out100, 1);

			System.out.println("completed...");

		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
