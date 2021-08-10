package sample;

import com.spire.doc.*;
import com.spire.doc.documents.DocumentObjectType;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.Field;
import sample.components.HintTextField;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Pattern;

public class SwingController {
    private JButton selectFilesBtn;
    private JToggleButton updateLinksBtn;
    private JToggleButton updateTemplatesBtn;
    private JToggleButton updateLogoBtn;
    private JTextField oldHyperlinkField;
    private JTextField newHyperlinkField;
    private JButton selectTemplateFileBtn;
    private JTextField oldTextField;
    private JTextField newTextField;
    private JButton addReplacementBtn;
    private JList<String> replacementList;
    private JButton removeReplacementBtn;
    private JButton selectDirectoryBtn;
    private JButton updateBtn;
    private JPanel mainPanel;


    private List<File> documents = new ArrayList<>();
    private DefaultListModel<String> replacementListModel;
    private File pictureFile;
    private File saveDirectory;

    private final  List<File> wordFiles = new ArrayList<>();
    private final List<File> pptFiles = new ArrayList<>();
    private final List<File> excelFiles = new ArrayList<>();

    public SwingController() {

        selectFilesBtn.addActionListener(new selectFilesListener());

        selectDirectoryBtn.addActionListener(new selectDirectoryListener());

        updateBtn.addActionListener(new updateListener(this));
    }

    public static void main(String[] args) {
        JFrame frame = new JFrame("Office Updater");
        frame.setContentPane(new SwingController().mainPanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }


    public static class updateListener implements ActionListener {

        private SwingController swingController;

        updateListener(SwingController controller) {
            swingController = controller;
        }

        @Override
        public void actionPerformed(ActionEvent e) {
            try {
                swingController.processUpdate();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }

    public class selectDirectoryListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            JFileChooser folderChooser = new JFileChooser(FileSystemView.getFileSystemView().getDefaultDirectory());
            folderChooser.setMultiSelectionEnabled(false);
            folderChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

            int result = folderChooser.showOpenDialog(null);
            if(result == JFileChooser.APPROVE_OPTION) {
                saveDirectory = folderChooser.getSelectedFile();
            }
        }
    }

    public class selectFilesListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            System.out.println("select files clicked");
            JFileChooser fileChooser = new JFileChooser(FileSystemView.getFileSystemView().getDefaultDirectory());
            fileChooser.setMultiSelectionEnabled(true);

            FileNameExtensionFilter officeFilter = new FileNameExtensionFilter("Microsoft Office Files",
                    "docx", "doc", "pptx", "ppt", "xlsx", "xlsm", "xlsb", "xltx", "xltm", "xls", "xlt", "xml");
            fileChooser.setFileFilter(officeFilter);
            fileChooser.setAcceptAllFileFilterUsed(false);

            int result = fileChooser.showOpenDialog(null);
            if(result == JFileChooser.APPROVE_OPTION) {
                documents.addAll(Arrays.asList(fileChooser.getSelectedFiles()));

                for (File document : documents) {
                    if(document.getName().contains(".doc")) {
                        wordFiles.add(document);
                    }
                }
            }
        }
    }



    private void createUIComponents() {
        oldHyperlinkField = new HintTextField("Old Hyperlink Directory");
        newHyperlinkField = new HintTextField("New Hyperlink Directory");
        oldTextField = new HintTextField("Old Text");
        newTextField = new HintTextField("New Text");
        replacementListModel = new DefaultListModel<>();
        replacementList = new JList<>(replacementListModel);
    }

    public void processUpdate() throws Exception {
        updateWordFiles(wordFiles);
    }

    private void updateWordFiles(List<File> files) {
        if(files.isEmpty()) {
            System.out.println("empty");
            return;
        }

        System.out.println("Updating Word Documents");
        List<Document> documents = new ArrayList<>();
        for(File file : files) {
            try {
                documents.add(new Document(file.getAbsolutePath()));
            } catch (Exception e) {
                System.out.println("\tLoad Document Exception: " + e.getMessage());
            }
        }

        for (int i = 0; i < documents.size(); i++) {

            Document document = documents.get(i);

            if (updateLinksBtn.isSelected()) {
                System.out.println("\tUpdating Hyperlinks");
                for (Section section : (Iterable<Section>) document.getSections()) {
                    for (Paragraph paragraph : (Iterable<Paragraph>) section.getParagraphs()) {
                        for (DocumentObject object : (Iterable<DocumentObject>) paragraph.getChildObjects()) {
                            if (object.getDocumentObjectType().equals(DocumentObjectType.Field)) { // check if object is a field
                                Field field = (Field) object;

                                // check if field is a hyperlink and matches the address to be replaced
                                if (field.getType().equals(FieldType.Field_Hyperlink) &&
                                        field.getCode().contains(oldHyperlinkField.getText())) {
                                    field.setCode(getNewHyperlink(field.getCode()));
                                }
                            }
                        }
                    }
                }
            }

            if (updateLogoBtn.isSelected()) {
                System.out.println("\tUpdating Logos");
                for (Section section : (Iterable<Section>) document.getSections()) {
                    //check headers/footers first
                    HeaderFooter header = section.getHeadersFooters().getHeader();
                    HeaderFooter footer = section.getHeadersFooters().getFooter();

                    updateLogosWord(header.getParagraphs());
                    updateLogosWord(footer.getParagraphs());
                    updateLogosWord(section.getParagraphs());
                }
            }


            if(!replacementListModel.isEmpty()) {
                for(int j = 0; j < replacementListModel.size(); j++) {
                    String[] replace = replacementListModel.get(j).split(" → ");
                    document.replace(Pattern.compile(replace[0]), replace[1]);
                }
            }

            String name = saveDirectory.getAbsolutePath() + "/" + files.get(i).getName().replace(".docx", "-test.docx");
            document.saveToFile(name, documents.get(i).getDetectedFormatType());
        }
    }

    private String getNewHyperlink(String original) {
        return original.replace(oldHyperlinkField.getText(), newHyperlinkField.getText());
    }

    private void updateLogosWord(Iterable<Paragraph> paragraphs) {
        for (Paragraph paragraph : paragraphs) {
            for (DocumentObject object : (Iterable<DocumentObject>) paragraph.getChildObjects()) {
                if (object instanceof DocPicture) {
                    System.out.println("docpicture");
                    DocPicture picture = (DocPicture) object;
                    if (picture.getAlternativeText().contains("har")) {
                        System.out.println("logo picture found");
                        picture.loadImage(pictureFile.getAbsolutePath());
                        picture.setAlternativeText("l3har_logo");
                    }
                }
            }
        }
    }
}
