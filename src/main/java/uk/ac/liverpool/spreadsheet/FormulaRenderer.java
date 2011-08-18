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

// Modified to render formulae in the XML format used by scone-fu

package uk.ac.liverpool.spreadsheet;

import java.util.Stack;

import org.apache.poi.hssf.util.AreaReference;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.formula.FormulaRenderingWorkbook;
import org.apache.poi.ss.formula.WorkbookDependentFormula;
import org.apache.poi.ss.formula.ptg.AbstractFunctionPtg;
import org.apache.poi.ss.formula.ptg.AreaPtg;
import org.apache.poi.ss.formula.ptg.AttrPtg;
import org.apache.poi.ss.formula.ptg.MemAreaPtg;
import org.apache.poi.ss.formula.ptg.MemErrPtg;
import org.apache.poi.ss.formula.ptg.MemFuncPtg;
import org.apache.poi.ss.formula.ptg.OperationPtg;
import org.apache.poi.ss.formula.ptg.ParenthesisPtg;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.RefPtg;

/**
 * Common logic for rendering formulas.<br/>
 *
 * Modified to render formulas in the XML format specified in SHAMAN WP9. Fabio Corubolo
 *
 * @author Josh Micich
 * 
 *  @author modifications Fabio Corubolo
 */
public class FormulaRenderer {

    /**
     * Static method to convert an array of {@link Ptg}s in RPN order
     * to a human readable string format in infix mode.
     * @param book  used for defined names and 3D references
     * @param ptgs  must not be <code>null</code>
     * @return a human readable String
     */
    public static String toFormulaString(FormulaRenderingWorkbook book, Ptg[] ptgs) {
        if (ptgs == null || ptgs.length == 0) {
            throw new IllegalArgumentException("ptgs must not be null");
        }
        Stack<String> stack = new Stack<String>();

        for (int i=0 ; i < ptgs.length; i++) {
            Ptg ptg = ptgs[i];
            // TODO - what about MemNoMemPtg?
            if(ptg instanceof MemAreaPtg || ptg instanceof MemFuncPtg || ptg instanceof MemErrPtg) {
                // marks the start of a list of area expressions which will be naturally combined
                // by their trailing operators (e.g. UnionPtg)
                // TODO - put comment and throw exception in toFormulaString() of these classes
                continue;
            }
            if (ptg instanceof ParenthesisPtg) {
                String contents = stack.pop();
                stack.push ("(" + contents + ")");
                continue;
            }
            if (ptg instanceof AttrPtg) {
                AttrPtg attrPtg = ((AttrPtg) ptg);
                if (attrPtg.isOptimizedIf() || attrPtg.isOptimizedChoose() || attrPtg.isSkip()) {
                    continue;
                }
                if (attrPtg.isSpace()) {
                    // POI currently doesn't render spaces in formulas
                    continue;
                    // but if it ever did, care must be taken:
                    // tAttrSpace comes *before* the operand it applies to, which may be consistent
                    // with how the formula text appears but is against the RPN ordering assumed here
                }
                if (attrPtg.isSemiVolatile()) {
                    // similar to tAttrSpace - RPN is violated
                    continue;
                }
                if (attrPtg.isSum()) {
                    String[] operands = getOperands(stack, attrPtg.getNumberOfOperands());
                    stack.push(attrPtg.toFormulaString(operands));
                    continue;
                }
                throw new RuntimeException("Unexpected tAttr: " + attrPtg.toString());
            }

            if (ptg instanceof WorkbookDependentFormula) {
                WorkbookDependentFormula optg = (WorkbookDependentFormula) ptg;
                stack.push(optg.toFormulaString(book));
                continue;
            }
            if (! (ptg instanceof OperationPtg)) {
                String s = "";
                if (ptg instanceof AreaPtg) {
                    AreaPtg a = (AreaPtg) ptg;
                    s = formatReferenceAsString(a);
                } else if (ptg instanceof RefPtg) {
                    RefPtg a = (RefPtg) ptg;
                    CellReference cr = new CellReference(a.getRow(), a.getColumn(), !a.isRowRelative(), !a.isColRelative());
                   s = "[." + cr.formatAsString() + "]";
                }   
                
                else s = ptg.toFormulaString();
                stack.push(s);
                continue;
            }

            OperationPtg o = (OperationPtg) ptg;
            
            String[] operands = getOperands(stack, o.getNumberOfOperands());
            if (o instanceof AbstractFunctionPtg) {
                AbstractFunctionPtg a = (AbstractFunctionPtg) o;
                stack.push(toFormulaString(a,operands));
            }else 
                stack.push(o.toFormulaString(operands));
        }
        if(stack.isEmpty()) {
            // inspection of the code above reveals that every stack.pop() is followed by a
            // stack.push(). So this is either an internal error or impossible.
            throw new IllegalStateException("Stack underflow");
        }
        String result = stack.pop();
        if(!stack.isEmpty()) {
            // Might be caused by some tokens like AttrPtg and Mem*Ptg, which really shouldn't
            // put anything on the stack
            throw new IllegalStateException("too much stuff left on the stack");
        }
        return result;
    }
    public static String toFormulaString(AbstractFunctionPtg a, String[] operands) {
        StringBuilder buf = new StringBuilder();

        if(a.isExternalFunction()) {
            buf.append(operands[0]); // first operand is actually the function name
            appendArgs(buf, 1, operands);
        } else {
            buf.append(a.getName());
            appendArgs(buf, 0, operands);
        }
        return buf.toString();
    }

    private static void appendArgs(StringBuilder buf, int firstArgIx, String[] operands) {
        buf.append('(');
        for (int i=firstArgIx;i<operands.length;i++) {
            if (i>firstArgIx) {
                buf.append(';');
            }
            buf.append(operands[i]);
        }
        buf.append(")");
    }
    
    protected static String formatReferenceAsString(AreaPtg a) {
        CellReference topLeft = new CellReference(a.getFirstRow(),a.getFirstColumn(),!a.isFirstRowRelative(),!a.isFirstColRelative());
        CellReference botRight = new CellReference(a.getLastRow(),a.getLastColumn(),!a.isLastRowRelative(),!a.isLastColRelative());

        if(AreaReference.isWholeColumnReference(topLeft, botRight)) {
                return (new AreaReference(topLeft, botRight)).formatAsString();
        }
        return "[." + topLeft.formatAsString() + ":" + "."+botRight.formatAsString()+"]";
}
    private static String[] getOperands(Stack<String> stack, int nOperands) {
        String[] operands = new String[nOperands];

        for (int j = nOperands-1; j >= 0; j--) { // reverse iteration because args were pushed in-order
            if(stack.isEmpty()) {
               String msg = "Too few arguments supplied to operation. Expected (" + nOperands
                    + ") operands but got (" + (nOperands - j - 1) + ")";
                throw new IllegalStateException(msg);
            }
            operands[j] = stack.pop();
        }
        return operands;
    }
}
