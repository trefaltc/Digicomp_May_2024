
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Variant;

public class Visa {

    public static void main(String[] args) {
        String sDir = "c:\\Users\\TN\\";
        String sInputDoc = sDir + "file_in.docx";
        String sOutputDoc = sDir + "file_out.docx";
        String sOldText = "[label:import:1]";
        String sNewText = "I am some horribly long sentence, so long that it may go to the next line if we keep going.";
        boolean tVisible = true;
        boolean tSaveOnExit = false;

        ActiveXComponent oWord = new ActiveXComponent("Word.Application");
        oWord.setProperty("Visible", tVisible);
        ActiveXComponent oDocuments = oWord.getPropertyAsComponent("Documents");
        ActiveXComponent oDocument = oDocuments.invokeGetComponent("Open", new Variant(sInputDoc));
        ActiveXComponent oSelection = oWord.getPropertyAsComponent("Selection");
        ActiveXComponent oFind = oSelection.getPropertyAsComponent("Find");
        oFind.setProperty("Text", sOldText);
        oFind.invoke("Execute");
        oSelection.setProperty("Text", sNewText);
        oSelection.invoke("MoveDown");
        oSelection.setProperty("Text", "\nSo we got the next line including BR.\n");
        ActiveXComponent oFont = oSelection.getPropertyAsComponent("Font");
        oFont.setProperty("Bold", "1");
        oFont.setProperty("Italic", "1");
        oFont.setProperty("Underline", "0");
        ActiveXComponent oWordBasic = oWord.getPropertyAsComponent("WordBasic");
        oWordBasic.invoke("FileSaveAs", sOutputDoc);
        oDocument.invoke("Close", tSaveOnExit);
        oWord.invoke("Quit", 0);

    }

}

