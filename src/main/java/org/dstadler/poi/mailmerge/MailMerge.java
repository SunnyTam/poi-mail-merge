package org.dstadler.poi.mailmerge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.logging.Logger;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.dstadler.commons.logging.jdk.LoggerFactory;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

/**
 * Simple application which performs a "mail-merge" of a Microsoft Word template
 * document which contains replacement templates in the form of ${name}, ${first-name}, ...
 * and an Microsoft Excel spreadsheet which contains a list of entries that are merged in.
 *
 * Call this application with parameters <word-template> <excel/csv-template> <output-file>
 *
 * The resulting document has all resulting documents concatenated.
 *
 * @author dominik.stadler
 *
 */
public class MailMerge {
	private static final Logger log = LoggerFactory.make();

	public static void main(String[] args) throws Exception {
		LoggerFactory.initLogging();

		if(args.length != 3) {
			throw new IllegalArgumentException("Usage: MailMerge <word-template> <excel/csv-template> <output-file>");
		}

		File wordTemplate = new File(args[0]);
		File excelFile = new File(args[1]);
		String outputFile = args[2];

		if(!wordTemplate.exists() || !wordTemplate.isFile()) {
			throw new IllegalArgumentException("Could not read Microsoft Word template " + wordTemplate);
		}
		if(!excelFile.exists() || !excelFile.isFile()) {
			throw new IllegalArgumentException("Could not read data file " + excelFile);
		}

		new MailMerge().merge(wordTemplate, excelFile, outputFile);
	}

	private void merge(File wordTemplate, File dataFile, String outputFile) throws Exception {
		log.info("Merging data from " + wordTemplate + " and " + dataFile + " into " + outputFile);

		// read the data-rows from the CSV or XLS(X) file
		Data data = new Data();
		data.read(dataFile);

		// now open the word file and apply the changes
		try (InputStream is = new FileInputStream(wordTemplate)) {
			try (HWPFDocument doc = new HWPFDocument(is)) {
				// apply the lines and concatenate the results into the document
				applyLines(data, doc);

			    log.info("Writing overall result to " + outputFile);
				try (OutputStream out = new FileOutputStream(outputFile)) {
			    	doc.write(out);
			    }
			}
		}
	}

	private void applyLines(Data dataIn, HWPFDocument doc){
		List<String> headers = dataIn.getHeaders();
		for(List<String> data : dataIn.getData()) {
			log.info("Applying to template: " + data);

			for(int fieldNr = 0;fieldNr < headers.size();fieldNr++) {
				String header = headers.get(fieldNr);
				String value = data.get(fieldNr);

				// ignore columns without headers as we cannot match them
				if(header == null) {
					continue;
				}

				// use empty string for data-cells that have no value
				if(value == null) {
					value = "";
				}

//				replaced = replaced.replace("${" + header + "}", value);
				replaceText(doc,"${" + header + "}", value);
			}

//			// check for missed replacements or formatting which interferes
//			if(replaced.contains("${")) {
//				log.warning("Still found template-marker after doing replacement: " +
//						StringUtils.abbreviate(StringUtils.substring(replaced, replaced.indexOf("${")), 200));
//			}
		}
	}

	private HWPFDocument replaceText(HWPFDocument doc, String findText, String replaceText){
		Range r1 = doc.getRange();

		for (int i = 0; i < r1.numSections(); ++i ) {
			Section s = r1.getSection(i);
			for (int x = 0; x < s.numParagraphs(); x++) {
				Paragraph p = s.getParagraph(x);
				for (int z = 0; z < p.numCharacterRuns(); z++) {
					CharacterRun run = p.getCharacterRun(z);
					String text = run.text();
					if(text.contains(findText)) {
						run.replaceText(findText, replaceText);
					}
				}
			}
		}
		return doc;
	}

	private void applyLines(Data dataIn, XWPFDocument doc) throws XmlException, IOException {
	    CTBody body = doc.getDocument().getBody();

	    XmlOptions optionsOuter = new XmlOptions();
	    optionsOuter.setSaveOuter();

	    // read the current full Body text
	    String srcString = body.xmlText();

	    // apply the replacements
	    boolean first = true;
	    List<String> headers = dataIn.getHeaders();
	    for(List<String> data : dataIn.getData()) {
	    	log.info("Applying to template: " + data);

	    	String replaced = srcString;
			for(int fieldNr = 0;fieldNr < headers.size();fieldNr++) {
	    		String header = headers.get(fieldNr);
	    		String value = data.get(fieldNr);

	    		// ignore columns without headers as we cannot match them
				if(header == null) {
	    			continue;
	    		}

				// use empty string for data-cells that have no value
				if(value == null) {
					value = "";
				}

				replaced = replaced.replace("${" + header + "}", value);
	    	}

			// check for missed replacements or formatting which interferes
			if(replaced.contains("${")) {
				log.warning("Still found template-marker after doing replacement: " +
						StringUtils.abbreviate(StringUtils.substring(replaced, replaced.indexOf("${")), 200));
			}

			appendBody(body, replaced, first);

			first = false;
	    }
	}

	private static void appendBody(CTBody src, String append, boolean first) throws XmlException {
	    XmlOptions optionsOuter = new XmlOptions();
	    optionsOuter.setSaveOuter();
	    String srcString = src.xmlText();
	    String prefix = srcString.substring(0,srcString.indexOf(">")+1);

	    final String mainPart;
	    // exclude template itself in first appending
	    if(first) {
	    	mainPart = "";
	    } else {
	    	mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
	    }

	    String sufix = srcString.substring( srcString.lastIndexOf("<") );
	    String addPart = append.substring(append.indexOf(">") + 1, append.lastIndexOf("<"));
	    CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
	    src.set(makeBody);
	}
}
