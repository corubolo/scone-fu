/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/*  * Author: Fabio Corubolo
 *  * Email: corubolo@gmail.com
 *  
 *******************************************************************************/
package uk.ac.liverpool.spreadsheet.example;

import java.io.File;
import java.io.FileInputStream;

import uk.ac.liverpool.spreadsheet.ToXML;



/**
 * 
 * @author Fabio Corubolo
 * 
 *         Test class for format conversion
 */

public class TestToML {

    public TestToML() {

    }

    /* Utility to test files */
    /* input: single file or folder */
    /* option -r to recurse */
    public static void main(String[] args) {
        boolean recurse = false;
        File d = null;
        if (args.length < 1)
            return;
        if (args[0].startsWith("-r")) {
            recurse = true;
            if (args.length < 2)
                return;
            d = new File(args[1]);
        } else
            d = new File(args[0]);
        TestToML.listFiles(d, recurse, 0);

    }

    public static void listFiles(File d, boolean recurse, int type) {

        if (d.isDirectory() && d.canRead() && d.exists()) {

            File[] dirList = d.listFiles();
            for (File element : dirList) {
                System.gc();
                if (!element.canRead()) {
                    System.out.println("Can't read " + element);
                }
                if (!element.exists()) {
                    System.out.println("Does not exist: " + element);
                }
                if (element.isDirectory()) {
                    if (recurse)
                        listFiles(element, recurse, type);
                    // loop again
                    continue;
                }
                try {
                    if (!element.getAbsolutePath().toLowerCase()
                            .endsWith("xls")
                            && !element.getAbsolutePath().toLowerCase()
                                    .endsWith("xlsx")) {
                         //System.out.println("Refused file: " + element);
                        continue;
                    }
                    try {
                        System.out.println("converting file: " + element);
                        ToXML toMl = ToXML.create(new FileInputStream(
                                element));
                        toMl.setEvaluateFormulae(false);
                        toMl.convert(element.getAbsolutePath() + ".xml");


                    } catch (Exception e) {
                        System.err.println("On file: " + element);
                        e.printStackTrace();
                    }

                } catch (Exception e) {
                    System.err.println("* On file: " + element);
                    e.printStackTrace();

                }
            }

        }

    }

}
