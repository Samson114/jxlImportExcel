package ONE;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ReadWriteExcelUtil {

	/**
	 * @param args
	 * @throws IOException 
	 * @throws WriteException 
	 * @throws RowsExceededException 
	 */
	public static void main(String[] args) throws RowsExceededException, WriteException, IOException {
//		String fileName = "d:" + File.separator + "students.xls";
//		System.out.println(ReadWriteExcelUtil.readExcel(fileName));
		String fileName1 = "d:" + File.separator + "abc.xls";
		ReadWriteExcelUtil.writeExcel(fileName1);
	}

	
	/**
	 * �у��݌���excel�ļ���
	 * 
	 * @param fileName
	 *            Ҫ������ļ������Q
	 * @throws IOException 
	 * @throws WriteException 
	 * @throws RowsExceededException 
	 */
	public static void writeExcel(String fileName) throws IOException, RowsExceededException, WriteException {
		WritableWorkbook wwb = null;
		
		// ����Ҫʹ��Workbook��Ĺ�����������һ����д��Ĺ�����(Workbook)����
		wwb = Workbook.createWorkbook(new File(fileName));
		
		if (wwb != null) {
			// ����һ����д��Ĺ�����
			// Workbook��createSheet������������������һ���ǹ���������ƣ��ڶ����ǹ������ڹ������е�λ��
			WritableSheet ws = wwb.createSheet("sheet1", 0);

			//�����ֶ���ӵ�Ԫ��
			int m=1;//��-1
			int n=4;//��-1
			String order_id="";
			String good_name="";
			String order_number="";
			String username="";
			
			String telephone="";
			String address="";
			String order_time="";
			String consumer_password="";
			
			Label label0 = new Label(1, 1, "�������");
			Label label1 = new Label(2, 1, "��Ʒ����");
			Label label2 = new Label(3, 1, "��Ʒ����");
			Label label3 = new Label(4, 1, "��ϵ��");
			
			Label label4 = new Label(5, 1, "��ϵ�绰");
			Label label5 = new Label(6, 1, "סַ");
			Label label6 = new Label(7, 1, "��������");
			Label label7 = new Label(8, 1, "������");
			
			ws.addCell(label0);
			ws.addCell(label1);
			ws.addCell(label2);
			ws.addCell(label3);
			ws.addCell(label4);
			ws.addCell(label5);
			ws.addCell(label6);
			ws.addCell(label7);
			
			  //���濪ʼ��ӵ�Ԫ��    
            for(int j=2;j<5;j++){    
				order_id="order_id"+n+"";
				good_name="good_name"+n+"";
				order_number="order_number"+n+"";
				username="username"+n+"";
				
				telephone="telephone"+n+"";
				address="address"+n+"";
				order_time="order_time"+n+"";
				consumer_password="consumer_password"+n+"";
				
				System.out.println("j="+j);
				
				
				Label label00 = new Label(1, j, order_id);//��һ��������  ��  �ڶ���������  ��
				Label label01 = new Label(2, j, good_name);
				Label label02 = new Label(3, j, order_number);
				Label label03 = new Label(4, j, username);
				
				Label label04 = new Label(5, j, telephone);
				Label label05 = new Label(6, j, address);
				Label label06 = new Label(7, j, order_time);
				Label label07 = new Label(8, j, consumer_password);
				// �����ɵĵ�Ԫ����ӵ���������
				
				ws.addCell(label00);
				ws.addCell(label01);
				ws.addCell(label02);
				ws.addCell(label03);
				
				ws.addCell(label04);
				ws.addCell(label05);
				ws.addCell(label06);
				ws.addCell(label07);
				
			}
			
			// ���ڴ���д���ļ���
			wwb.write();
			// �ر���Դ���ͷ��ڴ�
			wwb.close();
			
		}
			
	}

}
