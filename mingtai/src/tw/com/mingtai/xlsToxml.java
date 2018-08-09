package tw.com.mingtai;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;

/**
 * @author shouqing E-mail:shouqing777@gmail.com
 * @version 創建時間：2018年8月9日 下午2:53:44 類說明
 */
public class xlsToxml {
	public static void main(String[] args) throws IOException {

		ArrayList<String> arr = getFileList("D:\\liabatch\\CSV\\");

		for (String file : arr) {

			String line = null;

			int count = 0;

			FileReader fr = new FileReader("D:\\liabatch\\CSV\\" + file);

			FileWriter fw = new FileWriter("D:\\liabatch\\XML\\" + (file).substring(0, file.length() - 4) + ".txt");

			BufferedReader br = new BufferedReader(fr);

			fw.write("<ENDORSEMENTLIST>" + "\r\n");

			try {

				while ((line = br.readLine()).length() > 1) {

					count = count + 1;

					int comma1 = line.indexOf(',');
					int comma2 = line.indexOf(',', comma1 + 1);
					int comma3 = line.indexOf(',', comma2 + 1);
					int comma4 = line.indexOf(',', comma3 + 1);

					fw.write("<ENDORSEMENT>" + "\r\n");

					fw.write("<POLICYCOMPANY>" + line.substring(1, comma1 - 1) + "</POLICYCOMPANY>" + "\r\n");

					fw.write("<POLICYNUM>" + line.substring(comma1 + 2, comma2 - 1) + "</POLICYNUM>" + "\r\n");

					fw.write("<ENDORSEMENTCOMPANY>" + line.substring(comma2 + 2, comma3 - 1) + "</ENDORSEMENTCOMPANY>"
							+ "\r\n");

					fw.write("<ENDORSEMENTNUM>" + line.substring(comma3 + 2, comma4 - 1) + "</ENDORSEMENTNUM>" + "\r\n");

					fw.write("<BODY>"
							+ line.substring(comma4 + 2, line.length() - 1).replaceAll("\\x06", " ").replaceAll("\\x00"," ")
							+ "</BODY>" + "\r\n");

					fw.write("</ENDORSEMENT>" + "\r\n");

				}

				fw.write("</ENDORSEMENTLIST>");

			} catch (StringIndexOutOfBoundsException e) {

				fw.write("CSV內容第 : " + count + " 行發生問題，例外訊息為 " + e.toString() + " 字串長度有問題，請確認資料，謝謝。");

			}

			catch (Exception e) {

				fw.write("CSV內容第 : " + count + " 行發生問題，例外訊息為 " + e.toString());

			}

			fw.close();

			fr.close();

		}

	}

	public static ArrayList<String> getFileList(String folderPath) {

		ArrayList<String> arr = new ArrayList<String>();

		try {

			File folder = new File(folderPath);

			String[] list = folder.list();

			for (int i = 0; i < list.length; i++) {

				arr.add(list[i]);

			}

		} catch (Exception e) {

			System.out.println("'" + folderPath + "'此資料夾不存在");

		}

		return arr;

	}

}
