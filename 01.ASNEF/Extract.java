package extract;

import java.io.FileInputStream;
import java.io.File;
import java.io.StringReader;
import java.io.StringWriter;
import java.io.Writer;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Attr;
import org.w3c.dom.CDATASection;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hslf.util.SystemTimeUtils;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFStyles;



public class Extract {


	public static void main(String[] args) throws Exception {
		
		String DESTINATION_DIR="C:\\Users\\ddiaz.ext\\Desktop\\Contratos\\procesados\\";
		String FILE_DIR = "C:\\Users\\ddiaz.ext\\Desktop\\Contratos\\contratosdocx\\";
		File dir = new File(FILE_DIR);
		
		String[] list = dir.list();
		for (String file : list) {
			
			
		
		DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
	    DocumentBuilder docBuilder = docFactory.newDocumentBuilder();		
		Document doc = null;		
		
		System.err.println(FILE_DIR+file);
		
		XWPFDocument docx = new XWPFDocument(new FileInputStream(FILE_DIR+file));							
						
		// using XWPFWordExtractor Class
		XWPFWordExtractor we = new XWPFWordExtractor(docx);	
		
		
		
		
		
											
		doc = docBuilder.newDocument();
			
		//creamos la cabecera del fichero
		
		Element jasperHead = doc.createElement("jasperReport");
		jasperHead.setAttribute("xmlns", "http://jasperreports.sourceforge.net/jasperreports");
		jasperHead.setAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
		jasperHead.setAttribute("xsi:schemaLocation", "http://jasperreports.sourceforge.net/jasperreports http://jasperreports.sourceforge.net/xsd/jasperreport.xsd");	
		jasperHead.setAttribute("name", file);
		jasperHead.setAttribute("language", "groovy");
		jasperHead.setAttribute("pageWidth","595");
		jasperHead.setAttribute("pageHeight","842");
		jasperHead.setAttribute("columnWidth","555");
		jasperHead.setAttribute("leftMargin","20");
		jasperHead.setAttribute("rightMargin","20");
		jasperHead.setAttribute("topMargin","20");
		jasperHead.setAttribute("bottomMargin","20");
		jasperHead.setAttribute("uuid","c8d3a254-f72c-433b-808a-136805f147cc");
		doc.appendChild(jasperHead);	
		
		//Creamos las propiedades 
		Element properties1 = doc.createElement("property");
		properties1.setAttribute("name","ireport.zoom");
		properties1.setAttribute("value","1.0");
		jasperHead.appendChild(properties1);
		
		Element properties2 = doc.createElement("property");
		properties2.setAttribute("name","ireport.x");
		properties2.setAttribute("value","0");
		jasperHead.appendChild(properties2);
		
		Element properties3 = doc.createElement("property");
		properties3.setAttribute("name","ireport.y");
		properties3.setAttribute("value","0");
		jasperHead.appendChild(properties3);
		
		/*crear bandas
		Element background = doc.createElement("background");
			Element band = doc.createElement("band");
			band.setAttribute("splitType","Stretch");	
			background.appendChild(band);
		jasperHead.appendChild(background);
		*/
		
		//query por defecto
		Element queryString = doc.createElement("queryString");
		CDATASection cdata = doc.createCDATASection("select 1 from dual");
		queryString.appendChild(cdata);
		jasperHead.appendChild(queryString);

		Element field = doc.createElement("field");
		field.setAttribute("name","1");
		field.setAttribute("class","java.math.BigDecimal");
		jasperHead.appendChild(field);
		
		
		Element detail = doc.createElement("detail");
		jasperHead.appendChild(detail);
		
		
		
		
		//extraccion de parrafo completo 
		List<XWPFParagraph> parrafos= docx.getParagraphs();				
	    añadirParrafo(doc, parrafos, detail);
		
				
	    // extraccion del texto palabra por palabra
//		String[] stack = null;
//		stack = we.getText().split("\\n");				
//		añadirElemento(doc,stack,detail);
//		estamos comprobando que funciona el encriptado 
		
		
		
		
		// Introducimos los datos
		TransformerFactory transformerFactory = TransformerFactory
				.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
		transformer.setOutputProperty(OutputKeys.INDENT, "yes");
		DOMSource source = new DOMSource(doc);	
		
		//creamos una carpeta por cada archivo
		
		
		
		StreamResult result = new StreamResult(new File(DESTINATION_DIR+file.replace("docx", "jrxml")));

		
		
		// Output to console for testing
		// StreamResult result = new StreamResult(System.out);
		transformer.transform(source, result);
		//prettyPrint(doc);
		we.close();
		System.err.println("DOCUMENTO CREADO!");
		}
	}

	public static final void prettyPrint(Document xml) throws Exception {
		Transformer tf = TransformerFactory.newInstance().newTransformer();
		tf.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
		tf.setOutputProperty(OutputKeys.INDENT, "yes");
		Writer out = new StringWriter();
		tf.transform(new DOMSource(xml), new StreamResult(out));
		System.out.println(out.toString());
	}
	
	public static final Node  añadirElemento(Document doc,String[] stack,Node detail) throws Exception {
			
		int x=20;
		int y=0;	
		int height=12;
		int width=0;
		int maxWidth=535;
		int dynamicHeight=0;
		Element band=null;
		String value=null;		
		
		for (int i = 0; i < stack.length; i++) {	
			
			
			
			dynamicHeight=0;			
			value=stack[i];		
			//limpiamos caracteres extraños
			if(value.matches("\\t"))
			value=value.replaceAll("\\t", "");
			if(value.equals(" "))
				value=value.replace(" ", "");
								
			if (!value.equals(null) && !value.isEmpty()) {
				Attr attrdummy = null;											
				//maxima anchura campos				
				if(value.length()*6<maxWidth)
					width= (value.length()*6);
				else
					width=maxWidth;
						
										
				//creacion de detalles
				if(y%400==0 || y>400){
				    band = doc.createElement("band");
					band.setAttribute("height", "420");
					detail.appendChild(band);	
					y=0;	
				}
								
													
				// root elements
				Element staticText = doc.createElement("staticText");
				band.appendChild(staticText);				
				
				// distance elements
				Element reportElement = doc.createElement("reportElement");
				staticText.appendChild(reportElement);
				// set attribute to staff element Attr attr =
				// doc.createAttribute("id");attr.setValue("1");
				// x="308" y="87" width="71" height="20"				
				reportElement.setAttribute("uuid", "282dd817-2324-4978-944b-29896691c37a");				
				attrdummy = doc.createAttribute("x");
				attrdummy.setValue(String.valueOf(x));
				reportElement.setAttributeNode(attrdummy);
				attrdummy = doc.createAttribute("y");
								
					
				y+=height;
				height=12;
								
				attrdummy.setValue(String.valueOf(y));
				reportElement.setAttributeNode(attrdummy);
			
				
				attrdummy = doc.createAttribute("width");				
				attrdummy.setValue(String.valueOf(width));				
				reportElement.setAttributeNode(attrdummy);
				
				
				
				
				//altura dinamica de campos
				dynamicHeight=((value.length()/150)*height);
				if(dynamicHeight>height){				
					height= dynamicHeight+height;						
				}else if(value.length()>150){
					height=24;
				}
				band.setAttribute("height", String.valueOf(436+height));
				
				attrdummy = doc.createAttribute("height");
				attrdummy.setValue(String.valueOf(height));
				reportElement.setAttributeNode(attrdummy);
				
				
				Element property = doc.createElement("property");
				reportElement.appendChild(property);
				attrdummy = doc.createAttribute("name");
				attrdummy.setValue("ireport.layer");
				property.setAttributeNode(attrdummy);
				attrdummy = doc.createAttribute("value");
				attrdummy.setValue("1");
				property.setAttributeNode(attrdummy);

				Element textElement = doc.createElement("textElement");
				staticText.appendChild(textElement);
				Element font = doc.createElement("font");
				textElement.appendChild(font);
				attrdummy = doc.createAttribute("fontName");
				attrdummy.setValue("SansSerif");
				font.setAttributeNode(attrdummy);
				attrdummy = doc.createAttribute("size");
				attrdummy.setValue("8");
				font.setAttributeNode(attrdummy);

				Element text = doc.createElement("text");
				staticText.appendChild(text);
				CDATASection cdata = doc.createCDATASection(value);
				text.appendChild(cdata);
					
															
			}
		}

	 return detail;
	}
	
	public static final Node  añadirParrafo(Document doc,List<XWPFParagraph> parrafos,Node parrafo) throws Exception {
		
		int x=20;
		int y=0;	
		int height=12;
		int width=0;
		int maxWidth=535;
		int dynamicHeight=0;
		Element band=null;
		String value=null;
		
			for (XWPFParagraph xwpfParagraph : parrafos) {
				//XWPFStyles style=xwpfParagraph.getStyle();
				//xwpfParagraph.ge
				
		//	System.out.println("____"++"____"); 		
			
			ParagraphAlignment aling= xwpfParagraph.getAlignment();
			System.out.println(aling.toString()+"-----------       "+xwpfParagraph.getText());
		
			
			value=xwpfParagraph.getText();
								
			dynamicHeight=0;														
			if (!value.equals(null)) {
				Attr attrdummy = null;											
				//maxima anchura campos				
				if(value.length()*6<maxWidth)
					width= (value.length()*6);
				else
					width=maxWidth;
															
				//creacion de detalles
				if(y%400==0 || y>400){
				    band = doc.createElement("band");
					band.setAttribute("height", "420");
					parrafo.appendChild(band);	
					y=0;	
				}
																				
				// root elements
				Element staticText = doc.createElement("staticText");
				band.appendChild(staticText);				
				
				// distance elements
				Element reportElement = doc.createElement("reportElement");
				staticText.appendChild(reportElement);
				// set attribute to staff element Attr attr =
				// doc.createAttribute("id");attr.setValue("1");
				// x="308" y="87" width="71" height="20"				
				reportElement.setAttribute("uuid", "282dd817-2324-4978-944b-29896691c37a");				
				attrdummy = doc.createAttribute("x");
				attrdummy.setValue(String.valueOf(x));
				reportElement.setAttributeNode(attrdummy);
				attrdummy = doc.createAttribute("y");
												
				y+=height;
				height=12;
								
				attrdummy.setValue(String.valueOf(y));
				reportElement.setAttributeNode(attrdummy);
			
				
				attrdummy = doc.createAttribute("width");				
				attrdummy.setValue(String.valueOf(width));				
				reportElement.setAttributeNode(attrdummy);
				
											
				//altura dinamica de campos
				dynamicHeight=((value.length()/150)*height);
				if(dynamicHeight>height){				
					height= dynamicHeight+height;						
				}else if(value.length()>150){
					height=24;
				}
				band.setAttribute("height", String.valueOf(436+height));
				
				attrdummy = doc.createAttribute("height");
				attrdummy.setValue(String.valueOf(height));
				reportElement.setAttributeNode(attrdummy);
				
				
				Element property = doc.createElement("property");
				reportElement.appendChild(property);
				attrdummy = doc.createAttribute("name");
				attrdummy.setValue("ireport.layer");
				property.setAttributeNode(attrdummy);
				attrdummy = doc.createAttribute("value");
				attrdummy.setValue("1");
				property.setAttributeNode(attrdummy);

				Element textElement = doc.createElement("textElement");
				staticText.appendChild(textElement);
				Element font = doc.createElement("font");
				textElement.appendChild(font);
				attrdummy = doc.createAttribute("fontName");
				attrdummy.setValue("SansSerif");
				font.setAttributeNode(attrdummy);
				attrdummy = doc.createAttribute("size");
				attrdummy.setValue("8");
				font.setAttributeNode(attrdummy);

				Element text = doc.createElement("text");
				staticText.appendChild(text);
				CDATASection cdata = doc.createCDATASection(value);
				text.appendChild(cdata);
				
			}	
			
	}
	return parrafo;
	}
	
	
	

}