package pern.wang.bitmap2excelgrids;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;

public class Main {
    private static final short blockSize = 8;
    private static final int EXCEL_MAX_ITEM = 64000;

    public static void main(String[] args) throws Exception {
        //****************************************************
        //                    Read Image                     *
        //****************************************************
        File imgFile  = new File("D:\\3.jpg");
        if ( !imgFile.isFile() ) throw new Exception("Image file not exist!");
        BufferedImage bufImg = ImageIO.read(imgFile);

        int height = bufImg.getHeight();
        int width = bufImg.getWidth();
        int[][] orgImg = new int[height][width];
        for (int i = 0; i < height; i++) {
            for (int j = 0; j < width; j++) {
                orgImg[i][j] = bufImg.getRGB(j,i) & 0xFFFFFF;
            }
        }

        // get the zoom because you can just define up to 64000 style in a .xlsx
        int zoom = 1;
        if ( height*width>=EXCEL_MAX_ITEM ) zoom = (int)Math.ceil(Math.sqrt((double)height*(double)width/(double)EXCEL_MAX_ITEM));

        //****************************************************
        //                   Write Excel                     *
        //****************************************************
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("bitmap");

        for(int iRow=0; iRow<height/zoom; iRow++) {
            Row row = sheet.createRow(iRow);
            row.setHeightInPoints((int)(blockSize/1.3));
            for(int iCol=0; iCol<width/zoom; iCol++) {
                sheet.setColumnWidth(iCol,(int)35.7*blockSize);
                Cell cell = row.createCell(iCol);
                XSSFCellStyle style = wb.createCellStyle();
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                XSSFColor x = new XSSFColor();
                int ci = orgImg[iRow*zoom][iCol*zoom];
                byte[] cb = new byte[3];
                cb[2] = (byte)ci;
                cb[1] = (byte)(ci>>8);
                cb[0] = (byte)(ci>>16);
                x.setRGB(cb);
                style.setFillForegroundColor(x);
                cell.setCellStyle(style);
            }
        }

        FileOutputStream fileOut = new FileOutputStream("D:\\bigmap.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }
}
