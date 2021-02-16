package pl.jsolve.templ4docx.strategy;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

import pl.jsolve.templ4docx.insert.Insert;
import pl.jsolve.templ4docx.insert.TextInsert;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variable;

public class TextInsertStrategy implements InsertStrategy {

    @Override
    public void insert(Insert insert, Variable variable) {
        if (!(insert instanceof TextInsert)) {
            return;
        }
        if (!(variable instanceof TextVariable)) {
            return;
        }

        TextInsert textInsert = (TextInsert) insert;
        TextVariable textVariable = (TextVariable) variable;
        XWPFParagraph paragraph = textInsert.getParagraph();
        Map<XWPFRun, List<String>> runs = new HashMap<XWPFRun, List<String>>();
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);
            if (StringUtils.contains(text, textInsert.getKey().getKey())) 
            {
            	if (textVariable.getValue().contains( "\n" )) 
            	{
            		runs.put(run, new ArrayList<String>());
                    String[] stringsOnNewLines = textVariable.getValue().split( "\n" );
                    // For each additional line, create a new run. Add break return on previous.
                    for ( int i = 0; i < stringsOnNewLines.length; i++ ) {

                        // For every run except last one, add a break return.
                        String textForLine = stringsOnNewLines[i];
                        runs.get(run).add(textForLine);
                    }
                }
            	else
            	{
                    text = StringUtils.replace(text, textVariable.getKey(), textVariable.getValue());
                    run.setText(text, 0);
            	}
            }
        }
        addLines(paragraph, runs);
    }
    
    private void addLines(XWPFParagraph paragraph, Map<XWPFRun, List<String>> runs)
    {
    	for (Entry<XWPFRun, List<String>> run : runs.entrySet())
    	{
    		String[] lines = run.getValue().toArray(new String[0]);
    		for (int i = 0; i < lines.length; i++)
    		{
                if ( i < lines.length - 1 ) {
                	if (i == 0)
                	{
                		run.getKey().setText(lines[i], 0);
                	}
                	else
                	{
                		run.getKey().setText(lines[i]);
                	}
                    run.getKey().addBreak();
                }
                else {
                    XWPFRun newRun = paragraph.insertNewRun( i );
                    CTRPr rPr = newRun.getCTR().isSetRPr() ? newRun.getCTR().getRPr() : newRun.getCTR().addNewRPr();
                    rPr.set(run.getKey().getCTR().getRPr());
                    newRun.setText(lines[i]);
                }
    		}
    		
    	}
    }
}
