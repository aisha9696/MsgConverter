package kz.converter.msg;

import java.io.IOException;
import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import org.apache.poi.hsmf.exceptions.ChunkNotFoundException;

import org.apache.jempbox.xmp.XMPMetadata;
import org.apache.jempbox.xmp.pdfa.XMPSchemaPDFAId;
import org.apache.pdfbox.exceptions.COSVisitorException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentCatalog;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDMetadata;
import org.apache.pdfbox.pdmodel.edit.PDPageContentStream;




public class MsgConverter {
    public MsgConverter(String msgfile, String path) throws IOException, ChunkNotFoundException, COSVisitorException {
        MAPIMessage msg = new MAPIMessage(msgfile);



        String fromEmail = msg.getDisplayFrom();
        String subject = msg.getSubject();
        String body = msg.getTextBody();


        AttachmentChunks attachments[] = msg.getAttachmentFiles();
        if (attachments.length > 0) {
            for (AttachmentChunks a : attachments) {

                PDDocument doc = new PDDocument();
                PDPage page = new PDPage();
                doc.addPage(page);



                //ByteArrayInputStream fileIn = new ByteArrayInputStream(a.attachData.getValue());
				/*
				File f = new File(path, a.attachLongFileName.toString()); // output
				OutputStream fileOut = null;*/
                try {
					/*fileOut = new FileOutputStream(f);
					byte[] buffer = new byte[2048];
					int bNum = fileIn.read(buffer);
					while (bNum > 0) {
						fileOut.write(buffer);
						bNum = fileIn.read(buffer);
					}*/
                    PDPageContentStream contentStream = new PDPageContentStream(doc, page);


                    contentStream.beginText();
                    contentStream.drawString("From: " + fromEmail + " Subject " + subject + " Body " + body);
                    contentStream.drawString(a.attachData.toString());
                    contentStream.endText();



                    PDDocumentCatalog cat = doc.getDocumentCatalog();
                    PDMetadata metadata = new PDMetadata(doc);
                    cat.setMetadata(metadata);
                    // jempbox version
                    XMPMetadata xmp = new XMPMetadata();
                    XMPSchemaPDFAId pdfaid = new XMPSchemaPDFAId(xmp);
                    xmp.addSchema(pdfaid);
                    pdfaid.setConformance("B");
                    pdfaid.setPart(1);
                    pdfaid.setAbout("");


                    doc.save(path);

                } finally {
					/*try {
						if (fileIn != null) {
							fileIn.close();
						}
					} finally {
						if (fileOut != null) {
							fileOut.close();
						}
					}*/
                }
            }
        } else {

            System.out.println("No attachment");
        }
    }

    public static void main(String[] args) throws ChunkNotFoundException, COSVisitorException {
        if (args.length <= 0) {
            System.err.println("No files names provided");
        } else {
            try {
                new MsgConverter(args[0], args[1]);
            } catch (IOException e) {
                System.err.println("Could not process " + args[0] + ": " + e);
            }
        }
    }
}
