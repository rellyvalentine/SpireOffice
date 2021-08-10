package sample;

import com.spire.doc.*;
import com.spire.doc.FieldType;
import com.spire.doc.Section;
import com.spire.doc.documents.DocumentObjectType;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.Field;
import com.spire.presentation.*;
import com.spire.presentation.FileFormat;
import com.spire.presentation.collections.MasterSlideCollection;
import com.spire.xls.*;
import sample.components.HintTextField;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
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
    private JLabel filesLoadedLabel;
    private JPanel imageView;
    private JLabel updateCompleteLabel;


    private List<File> documents = new ArrayList<>();
    private DefaultListModel<String> replacementListModel;
    private File pictureFile;
    private File templateFile;
    private File saveDirectory;

    private final  List<File> wordFiles = new ArrayList<>();
    private final List<File> pptFiles = new ArrayList<>();
    private final List<File> excelFiles = new ArrayList<>();

    public SwingController() {

        selectFilesBtn.addActionListener(new selectFilesListener());

        selectDirectoryBtn.addActionListener(new selectDirectoryListener());

        selectTemplateFileBtn.addActionListener(new selectTemplateListener());

        imageView.addMouseListener(new selectImageListener(this));

        addReplacementBtn.addActionListener(new addReplacementListener());

        removeReplacementBtn.addActionListener(new removeReplacementListener());

        updateBtn.addActionListener(new updateListener(this));
    }

    public static void main(String[] args) {
        JFrame frame = new JFrame("Office Updater");
        frame.setContentPane(new SwingController().mainPanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }

    /* ============ ACTION LISTENERS ============*/

    public static class selectImageListener extends MouseAdapter {
        private SwingController swingController;

        selectImageListener(SwingController controller) {
            swingController = controller;
        }

        @Override
        public void mouseClicked(MouseEvent e) {
            System.out.println("clicked imageview");
            JFileChooser fileChooser = new JFileChooser(FileSystemView.getFileSystemView().getDefaultDirectory());
            FileNameExtensionFilter officeFilter = new FileNameExtensionFilter("Image Files",
                    "jpg", "jpeg", "jpe", "eps", "png", "svg", "svgz");
            fileChooser.setFileFilter(officeFilter);
            fileChooser.setMultiSelectionEnabled(false);
            fileChooser.setAcceptAllFileFilterUsed(false);

            int result = fileChooser.showOpenDialog(null);
            if(result == JFileChooser.APPROVE_OPTION) {
                swingController.pictureFile = fileChooser.getSelectedFile();
                swingController.addLogo();
            }
        }
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

    public class selectTemplateListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            System.out.println("select template clicked");
            JFileChooser fileChooser = new JFileChooser(FileSystemView.getFileSystemView().getDefaultDirectory());
            FileNameExtensionFilter templateFilter = new FileNameExtensionFilter("PowerPoint Template Files", "potx");
            fileChooser.setFileFilter(templateFilter);
            fileChooser.setMultiSelectionEnabled(false);
            fileChooser.setAcceptAllFileFilterUsed(false);

            int result = fileChooser.showOpenDialog(null);
            if(result == JFileChooser.APPROVE_OPTION) {
                templateFile = fileChooser.getSelectedFile();
            }
        }
    }

    public class selectFilesListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            System.out.println("select files clicked");
            JFileChooser fileChooser = new JFileChooser(FileSystemView.getFileSystemView().getDefaultDirectory());

            FileNameExtensionFilter officeFilter = new FileNameExtensionFilter("Microsoft Office Files",
                    "docx", "doc", "pptx", "ppt", "xlsx", "xlsm", "xlsb", "xltx", "xltm", "xls", "xlt", "xml");
            fileChooser.setFileFilter(officeFilter);
            fileChooser.setMultiSelectionEnabled(true);
            fileChooser.setAcceptAllFileFilterUsed(false);

            int result = fileChooser.showOpenDialog(null);
            if(result == JFileChooser.APPROVE_OPTION) {
                documents.addAll(Arrays.asList(fileChooser.getSelectedFiles()));
                for (File document : documents) {
                    if(document.getName().contains(".doc")) {
                        wordFiles.add(document);
                    }
                }
                for(File document : documents) {
                    if(document.getName().contains(".ppt")) {
                        pptFiles.add(document);
                    }
                }
                for(File document : documents) {
                    if(document.getName().contains(".x")) {
                        excelFiles.add(document);
                    }
                }
                updateCompleteLabel.setVisible(false);
            }
            filesLoadedLabel.setVisible(true);
        }
    }

    public class addReplacementListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            String oldText = oldTextField.getText();
            String newText = newTextField.getText();
            if(!oldText.equals("") && !oldText.equals("Old Text") && !newText.equals("") && !newText.equals("New Text")) {
                String replacement = oldText + " → " + newText;
                replacementListModel.addElement(replacement);
                oldTextField.setText("Old Text");
                newTextField.setText("New Text");
            }
        }
    }

    public class removeReplacementListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            replacementListModel.remove(replacementList.getSelectedIndex());
        }
    }

    private void createUIComponents() {
        oldHyperlinkField = new HintTextField("Old Hyperlink Directory");
        newHyperlinkField = new HintTextField("New Hyperlink Directory");
        oldTextField = new HintTextField("Old Text");
        newTextField = new HintTextField("New Text");
        replacementListModel = new DefaultListModel<>();
        replacementList = new JList<>(replacementListModel);
        imageView = new JPanel();
    }

    /* ============ IMPLEMENTATION ============*/

    public void processUpdate() throws Exception {
        updateWordFiles(wordFiles);
        updatePowerPointFiles(pptFiles);
        updateExcelFiles(excelFiles);
        resetDisplay();
        updateCompleteLabel.setVisible(true);

    }

    private void resetDisplay() {
        filesLoadedLabel.setVisible(false);

        updateLinksBtn.setSelected(false);
        updateLogoBtn.setSelected(false);
        updateTemplatesBtn.setSelected(false);

        oldHyperlinkField.setText("Old Hyperlink Directory");
        newHyperlinkField.setText("New Hyperlink Directory");
        oldTextField.setText("Old Text");
        newTextField.setText("New Text");

        imageView.removeAll();

        replacementListModel.removeAllElements();
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

    private void updatePowerPointFiles(List<File> files) throws Exception {
        if (files.isEmpty())
            return;
        System.out.println("Updating PowerPoint Files...");
        List<Presentation> presentations = new ArrayList<>();
        for(File file : files) {
            try {
                Presentation presentation = new Presentation();
                presentation.loadFromFile(file.getAbsolutePath());
                presentations.add(presentation);
            } catch (Exception e) {
                System.out.println("Presentation Load Error: " + e.getMessage());
            }
        }

        for (int i = 0; i < presentations.size(); i++) {

            Presentation presentation = presentations.get(i);

            // update PPT Hyperlinks & words
            if (updateLinksBtn.isSelected()) {
                System.out.println("\tUpdating Hyperlinks");
                for (Object slide : presentation.getSlides()) {
                    if (slide instanceof ISlide) {
                        for (Object shape : ((ISlide) slide).getShapes()) {
                            if (shape instanceof IShape) {
                                if (((IShape) shape).getClick() != null) { //replace hyperlink embedded in image/shape
                                    if (((IShape) shape).getClick().getAddress() != null) {
                                        String hyperlink = ((IShape) shape).getClick().getAddress();
                                        if (hyperlink.contains(oldHyperlinkField.getText())) {
                                            ((IShape) shape).getClick().setAddress(getNewHyperlink(hyperlink));
                                        }

                                    }
                                }
                                else { // replace hyperlink text and words
                                    if(shape instanceof IAutoShape) {
                                        for(Object paragraph : ((IAutoShape) shape).getTextFrame().getParagraphs()) {
                                            ParagraphEx pe = (ParagraphEx)paragraph;
                                            for(Object textPortion : pe.getTextRanges()) {
                                                PortionEx te = (PortionEx) textPortion;

                                                if(te.getClickAction() != null) { // change hyperlink address
                                                    System.out.println("hyperlink name: "+((IShape) shape).getName());
                                                    String hyperlink = te.getClickAction().getAddress();
                                                    if(hyperlink != null && hyperlink.contains(oldHyperlinkField.getText())) {
                                                        te.getClickAction().setAddress(getNewHyperlink(hyperlink));
                                                    }
                                                }

                                                if(!replacementListModel.isEmpty()) {
                                                    for (int j = 0; j < replacementListModel.size(); j++) {
                                                        String[] replace = replacementListModel.get(j).split(" → ");
                                                        if (te.getText().contains(replace[0])) {
                                                            te.setText(te.getText().replace(replace[0], replace[1]));
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // update PPT Templates
            if (updateTemplatesBtn.isSelected() && templateFile != null) {
                System.out.println("\tUpdating Templates...");
                Presentation template = new Presentation();
                template.loadFromFile(templateFile.getAbsolutePath());
                MasterSlideCollection masterSlides = presentation.getMasters();
                masterSlides.appendSlide(template.getMasters().get(0));

                // set title layouts
                ISlide firstSlide = presentation.getSlides().get(0);

                firstSlide
                        .setLayout(masterSlides.get(masterSlides.getCount() - 1)   // get replacement master slide
                                .getLayouts().get(0));                          // get title layout

                if (firstSlide.getTitle().equals("")) { // title is contained within a textbox
                    Object shape = firstSlide.getShapes().get(1);
                    if (shape instanceof IAutoShape) {
                        IAutoShape titleContainer = (IAutoShape) shape;
                        firstSlide.setTitle(titleContainer.getTextFrame()
                                .getParagraphs().get(0).getText());

                        firstSlide.getShapes().removeAt(1); // remove duplicate text
                    } else {
                        System.out.println("Title could not be found");
                    }

                    for (int j = 0; j < firstSlide.getShapes().size(); j++) {
                        shape = firstSlide.getShapes().get(j);
                        if (shape instanceof SlidePicture) {
                            firstSlide.getShapes().removeAt(j);
                        }
                    }
                }

                // set all other slides layout
                for (int j = 1; j < presentation.getSlides().getCount(); j++) {
                    ISlide slide = presentation.getSlides().get(j);
                    List<Float> positions = new ArrayList<>();

                    // store positions for each shape
                    for (Object shape : slide.getShapes()) {
                        if (shape instanceof IAutoShape) {
                            positions.add(((IAutoShape) shape).getLeft());
                        } else if (shape instanceof IShape) {
                            positions.add(((IShape) shape).getLeft());
                        }
                    }
                    slide.setLayout(masterSlides.get(masterSlides.getCount() - 1)
                            .getLayouts().get(3));

                    // re-apply the positions
                    for (int k = 0; k < positions.size(); k++) {
                        slide.getShapes().get(k)
                                .setLeft(positions.get(k));
                    }
                }
            }

            String name = saveDirectory.getAbsolutePath() + "/" + files.get(i).getName().replace(".pptx", "-test.pptx");
            presentation.saveToFile(name, com.spire.presentation.FileFormat.PPTX_2013);
        }
    }

    private void updateExcelFiles(List<File> files) {
        if (files.isEmpty())
            return;

        System.out.println("Updating Excel Files...");
        List<Workbook> workbooks = new ArrayList<>();
        for(File file : files) {
            try {
                Workbook wb = new Workbook();
                wb.loadFromFile(file.getAbsolutePath());
                workbooks.add(wb);
            } catch (Exception e) {
                System.out.println("Error loading Excel Workbook: " + e.getMessage());
            }
        }

        for (int i = 0; i < workbooks.size(); i++) {
            Workbook wb = workbooks.get(i);

            for(Object ws : wb.getWorksheets()) {
                if(ws instanceof Worksheet) {

                    if(updateLinksBtn.isSelected()) { // update hyperlink
                        System.out.println("\tUpdating Hyperlinks");
                        for (Object hyperLink : ((Worksheet) ws).getHyperLinks()) {
                            if (hyperLink instanceof HyperLink) {
                                if (((HyperLink) hyperLink).getAddress().contains(oldHyperlinkField.getText())) {
                                    ((HyperLink) hyperLink).setAddress(getNewHyperlink(((HyperLink) hyperLink).getAddress()));
                                }
                            }
                        }
                    }

                    if(updateLogoBtn.isSelected()) {
                        for (Object picture : ((Worksheet) ws).getPictures()) {
                            if (picture instanceof ExcelPicture) {
                                if (((ExcelPicture) picture).getAlternativeText().contains("har")) {
                                    BufferedImage img;
                                    try {
                                        img = ImageIO.read(pictureFile);
                                        ((ExcelPicture) picture).setPicture(img);
                                        ((ExcelPicture) picture).setAlternativeText("l3har_logo");
                                    } catch (IOException e) {
                                        e.printStackTrace();
                                    }
                                }
                            }
                        }
                    }

                    if(!replacementListModel.isEmpty()) {
                        for (int j = 0; j < replacementListModel.size(); j++) {
                            String[] replace = replacementListModel.get(j).split(" → ");
                            for(CellRange range : ((Worksheet) ws).findAllString(replace[0], true, true)) {
                                range.setText(range.getText().replace(replace[0], replace[1]));
                            }
                        }
                    }
                }
            }
            String name = saveDirectory.getAbsolutePath() + "/" + files.get(i).getName().replace(".xlsx", "-test.xlsx");
            wb.saveToFile(name);
        }
    }

    public void addLogo() {
        if(pictureFile != null) {
                    try {
                        BufferedImage logo = ImageIO.read(pictureFile);
                        JLabel picLabel =  new JLabel(new ImageIcon(logo.getScaledInstance(200, 50, Image.SCALE_DEFAULT)));
                        picLabel.setVisible(true);
//                        System.out.println(picLabel.getX() + " "+ picLabel.getY());
                        imageView.add(picLabel);
                        // not sure why but image literally doesn't display unless i do this
                        updateCompleteLabel.setVisible(true);
                        updateCompleteLabel.setVisible(false);
                    } catch (IOException ex) {
                        ex.printStackTrace();
                    }
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
