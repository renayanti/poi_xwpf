/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package myproject.xwpf.app;

import com.microsoft.schemas.vml.CTGroup;
import com.microsoft.schemas.vml.CTShape;
import com.microsoft.schemas.vml.CTShapetype;
import com.microsoft.schemas.vml.STTrueFalse;
import java.awt.EventQueue;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSchemeColor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

/**
 *
 * @author x
 */
public final class Process extends JPanel {
	FileInputStream fc;
	XWPFDocument document;
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			@Override
			public void run() {
				try {
					UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
				} catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedOperationException ex) {
					ex.printStackTrace();
				} catch (UnsupportedLookAndFeelException ex) {
					Logger.getLogger(Process.class.getName()).log(Level.SEVERE, null, ex);
				}

				JFrame frame = new JFrame("test");
				frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
				frame.add(new MainPanel());
				frame.pack();
				frame.setLocationRelativeTo(null);

				frame.setVisible(true);
				frame.getComponents();
			}

		});
	}

	Process(FileInputStream in) throws FileNotFoundException,IOException{
		fc =in;
		try {
			document  = new XWPFDocument(fc);
			
			List<XWPFParagraph> paragraph;
			paragraph = document.getParagraphs();
			int i = 100;
			for (XWPFParagraph para : paragraph){
//				int Tab = para.getIndentationHanging();
//				System.out.println(Tab);
//				 CTGroup ctGroup = CTGroup.Factory.newInstance();
//				 CTShape ctShape = ctGroup.addNewShape();
//				ctShape.setStyle("width:80pt;height:24pt");
//				CTP ctp = CTP.Factory.newInstance();
				CTP ctp = para.getCTP();
				CTShape zz = createShape(ctp,i);
				
				document.setParagraph(zz, i);
				i++;
			}
			document.write(new FileOutputStream ("/home/x/Documents/temp/xwpfdoc2.docx"));
		} catch (Exception e) {
			Logger.getLogger(Process.class.getName()).log(Level.SEVERE, null, e);
		}
		
		
	}
	
	protected CTShape createShape(CTP ctp,int  idx) {
//		BigInteger zz  = ctp.getPPr().getRPr().getKern().getVal();
//		System.out.println(zz);
		CTGroup group = CTGroup.Factory.newInstance();
//		CTShapetype shapetype = group.addNewShapetype();
//		shapetype.setId("_x0000_t136");
//		shapetype.setCoordsize("1600,21600");
//		shapetype.setSpt(136);
		CTShape shape = group.addNewShape();
		shape.setId("PowerPlusWaterMarkObject" + idx);
		shape.setSpid("_x0000_s102" + (400 + idx));
		shape.setType("#_x0000_t136");
		shape.setStyle("position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin");
		shape.setWrapcoords("616 5068 390 16297 39 16921 -39 17155 7265 17545 7186 17467 -39 17467 18904 17467 10507 17467 8710 17545 18904 17077 18787 16843 18358 16297 18279 12554 19178 12476 20701 11774 20779 11228 21131 10059 21248 8811 21248 7563 20975 6316 20935 5380 19490 5146 14022 5068 2616 5068");
		shape.setFillcolor("black");
		shape.setStroked(STTrueFalse.FALSE);
		CTPPr rPr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
//		rPr.set(shape);
		return shape;
	}

}
