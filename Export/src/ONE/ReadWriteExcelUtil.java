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
	 * 把內容寫入excel文件中
	 * 
	 * @param fileName
	 *            要寫入的文件的名稱
	 * @throws IOException 
	 * @throws WriteException 
	 * @throws RowsExceededException 
	 */
	public static void writeExcel(String fileName) throws IOException, RowsExceededException, WriteException {
		WritableWorkbook wwb = null;
		
		// 首先要使用Workbook类的工厂方法创建一个可写入的工作薄(Workbook)对象
		wwb = Workbook.createWorkbook(new File(fileName));
		
		if (wwb != null) {
			// 创建一个可写入的工作表
			// Workbook的createSheet方法有两个参数，第一个是工作表的名称，第二个是工作表在工作薄中的位置
			WritableSheet ws = wwb.createSheet("sheet1", 0);

			//下面手动添加单元格
			int m=1;//列-1
			int n=4;//行-1
			String order_id="";
			String good_name="";
			String order_number="";
			String username="";
			
			String telephone="";
			String address="";
			String order_time="";
			String consumer_password="";
			
			Label label0 = new Label(1, 1, "订单编号");
			Label label1 = new Label(2, 1, "商品名称");
			Label label2 = new Label(3, 1, "商品数量");
			Label label3 = new Label(4, 1, "联系人");
			
			Label label4 = new Label(5, 1, "联系电话");
			Label label5 = new Label(6, 1, "住址");
			Label label6 = new Label(7, 1, "订单日期");
			Label label7 = new Label(8, 1, "消费码");
			
			ws.addCell(label0);
			ws.addCell(label1);
			ws.addCell(label2);
			ws.addCell(label3);
			ws.addCell(label4);
			ws.addCell(label5);
			ws.addCell(label6);
			ws.addCell(label7);
			
			  //下面开始添加单元格    
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
				
				
				Label label00 = new Label(1, j, order_id);//第一个参数是  列  第二个参数是  行
				Label label01 = new Label(2, j, good_name);
				Label label02 = new Label(3, j, order_number);
				Label label03 = new Label(4, j, username);
				
				Label label04 = new Label(5, j, telephone);
				Label label05 = new Label(6, j, address);
				Label label06 = new Label(7, j, order_time);
				Label label07 = new Label(8, j, consumer_password);
				// 将生成的单元格添加到工作表中
				
				ws.addCell(label00);
				ws.addCell(label01);
				ws.addCell(label02);
				ws.addCell(label03);
				
				ws.addCell(label04);
				ws.addCell(label05);
				ws.addCell(label06);
				ws.addCell(label07);
				
			}
			
			// 从内存中写入文件中
			wwb.write();
			// 关闭资源，释放内存
			wwb.close();
			
		}
			
	}

}
