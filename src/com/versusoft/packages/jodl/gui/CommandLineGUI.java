/**
 *  odt2daisy - OpenDocument to DAISY XML/Audio
 *
 *  (c) Copyright 2008 - 2009 by Vincent Spiewak, All Rights Reserved.
 *
 *  This program is free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Lesser Public License as published by
 *  the Free Software Foundation; either version 3 of the License, or 
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU Lesser General Public License for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public License along
 *  with this program; if not, write to the Free Software Foundation, Inc.,
 *  51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.	
 */
package com.versusoft.packages.jodl.gui;

import com.versusoft.packages.jodl.*;
import java.io.File;
import java.io.IOException;
import java.util.logging.FileHandler;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.xml.sax.SAXException;

/**
 * Command Line Interface
 * 
 * @author Vincent Spiewak <vspiewak at gmail dot com>
 */
public class CommandLineGUI {

    private static final String LOG_FILENAME_PATTERN = "%t/jodl.log";
    private static final Logger logger = Logger.getLogger("com.versusoft.packages.jodl.odtutils");

    public static void main(String args[]) throws SAXException, IOException {

        Handler fh = new FileHandler(LOG_FILENAME_PATTERN);
        fh.setFormatter(new SimpleFormatter());

        //removeAllLoggersHandlers(Logger.getLogger(""));
        Logger.getLogger("").addHandler(fh);
        Logger.getLogger("").setLevel(Level.FINEST);



        Options options = new Options();

        Option option1 = new Option("in", "ODT file (required)");
        option1.setRequired(true);
        option1.setArgs(1);

        Option option2 = new Option("out", "Output file (required)");
        option2.setRequired(false);
        option2.setArgs(1);

        Option option3 = new Option("pic", "extract pics");
        option3.setRequired(false);
        option3.setArgs(1);

        Option option4 = new Option("page", "enable pagination processing");
        option4.setRequired(false);
        option4.setArgs(0);

        options.addOption(option1);
        options.addOption(option2);
        options.addOption(option3);
        options.addOption(option4);

        CommandLineParser parser = new BasicParser();
        CommandLine cmd = null;

        try {
            cmd = parser.parse(options, args);
        } catch (ParseException e) {
            printHelp();
            return;
        }

        if (cmd.hasOption("help")) {
            printHelp();
            return;
        }

        File outFile = new File(cmd.getOptionValue("out"));

        OdtUtils utils = new OdtUtils();


        utils.open(cmd.getOptionValue("in"));
        utils.correctionStep();
        utils.saveXML(outFile.getAbsolutePath());

        if (cmd.hasOption("page")) {
            try {

                OdtUtils.paginationProcessing(outFile.getAbsolutePath());

            } catch (ParserConfigurationException ex) {
                logger.log(Level.SEVERE, null, ex);
            } catch (SAXException ex) {
                logger.log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                logger.log(Level.SEVERE, null, ex);
            } catch (TransformerConfigurationException ex) {
                logger.log(Level.SEVERE, null, ex);
            } catch (TransformerException ex) {
                logger.log(Level.SEVERE, null, ex);
            }
        }

        if (cmd.hasOption("pic")) {

            String imageDir = cmd.getOptionValue("pic");
            if (!imageDir.endsWith("/")) {
                imageDir += "/";
            }

            try {

                String basedir = new File(cmd.getOptionValue("out")).getParent().toString() + System.getProperty("file.separator");
                OdtUtils.extractAndNormalizedEmbedPictures(cmd.getOptionValue("out"), cmd.getOptionValue("in"), basedir, imageDir);
            } catch (SAXException ex) {
                logger.log(Level.SEVERE, null, ex);
            } catch (ParserConfigurationException ex) {
                logger.log(Level.SEVERE, null, ex);
            } catch (TransformerConfigurationException ex) {
                logger.log(Level.SEVERE, null, ex);
            } catch (TransformerException ex) {
                logger.log(Level.SEVERE, null, ex);
            }
        }

    }

    static void printHelp() {
        System.out.println("Usage: -odt odtfile -xslt xsltfile");
        System.out.println("");
        System.out.println("Required Params:");
        System.out.println("-in     ODT file (required)");
        System.out.println("-out    Output file (required)");
        System.out.println("");
        System.out.println("Optionals Params:");
        System.out.println("-pic      extract pictures");
        System.out.println("-page     enable pagination");
        System.out.println("");
        System.out.println("Examples:");
        System.out.println("java -jar JODL -in ./input.odt -out ./output.xml");
        System.out.println("java -jar JODL -in ./input.odt -out ./output.xml -pic images/");
        System.out.println("java -jar JODL -in ./input.odt -out ./output.xml -pic images/ -page");
        System.out.println("");
        System.out.println("(C) Copyright 2008  Vincent Spiewak");

    }
}
