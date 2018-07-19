package pdf;

import com.lowagie.text.DocumentException;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfReader;

import java.io.IOException;

public class Test {
    private static String rootpath = "C:\\Users\\Harley\\Desktop\\1218\\";
    private static String pdfFILE_DI = rootpath + "1.pdf";
    private static String pdfFILE_ou = rootpath + "240_2.pdf";
    private static String pdfimg = rootpath + "240_2.bmp";

    public static void main(String[] args) {

        test();
    }

    public static void test() {
        PdfReader reader = null;
        try {
            reader = new PdfReader(pdfFILE_DI);
        } catch (IOException e) {
            e.printStackTrace();
        }
        Rectangle a = reader.getPageSize(1);
        float sizex = a.getWidth();
        float sizey = a.getHeight();
        System.out.println("背景点数大小" + sizex + "  " + sizey);
        reader.close();
        //定义插入位置
        float x = (float) 0, y = (float) 0, X = sizex, Y = sizey;
        // float x = (float) 100, y = (float) 100, X = (float) 61, Y = (float) 61;
        try {
            InsertPdf.insert(pdfFILE_DI, pdfFILE_ou, pdfimg, x, y, X, Y);

        } catch (DocumentException | IOException e) {
            e.printStackTrace();
        }
    }
}

