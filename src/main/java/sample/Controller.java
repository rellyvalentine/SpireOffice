package sample;

import javax.swing.*;
import java.awt.*;

public class Controller extends JFrame {
    public Controller() {
        JPanel panel = new JPanel();
        panel.setLayout(new GridLayout(4, 1));
        JLabel l1, l2, l3, l4;
        l1 = new JLabel("top row");
        l2 = new JLabel("middle row");
        l3 = new JLabel("l3");
        l4 = new JLabel("l4");

        panel.add(l1);
        panel.add(l2);
        panel.add(l3);
        panel.add(l4);

        add(panel);
        setVisible(true);
        this.setSize(800, 700);

    }
}

//package sample;
//
//import com.spire.doc.*;
//import com.spire.doc.FieldType;
//import com.spire.doc.Section;
//import com.spire.doc.documents.DocumentObjectType;
//import com.spire.doc.documents.Paragraph;
//import com.spire.doc.fields.DocPicture;
//import com.spire.doc.fields.Field;
//import com.spire.presentation.*;
//import com.spire.presentation.FileFormat;
//import com.spire.presentation.collections.MasterSlideCollection;
//import com.spire.xls.*;
//
//public class Controller {
//
//    private final FileChooser imageChooser = new FileChooser();
//    private final FileChooser templateChooser = new FileChooser();
//    private final FileChooser docChooser = new FileChooser();
////    private final DirectoryChooser directoryChooser = new DirectoryChooser();
//
//    private File pictureFile;
//    private File templateFile;
////    private File saveDirectory;
//
//    @FXML
//    private GridPane leftGrid;
//
//    @FXML
//    private GridPane rightGrid;
//
//    @FXML
//    private Label loadedLabel;
//
//    @FXML
//    private ToggleButton updateLinksBtn;
//
//    @FXML
//    private ToggleButton updateLogoBtn;
//
//    @FXML
//    private ToggleButton updateTemplateBtn;
//
//    @FXML
//    private Button selectTemplateFile;
//
//    @FXML
//    private VBox updateLinksGrid;
//
//    @FXML
//    private GridPane updateLogoGrid;
//
//    @FXML
//    private TextField oldHyperlinkField;
//
//    @FXML
//    private TextField newHyperlinkField;
//
//    @FXML
//    private TextField oldText;
//
//    @FXML
//    private TextField newText;
//
//    @FXML
//    private ImageView imageView;
//
//    @FXML
//    private ListView<Label> replacementList;
//
////    @FXML
////    private Button selectDirectory;
//
//    @FXML
//    private Button updateBtn;
//
//    @FXML
//    private Label updateCompleteLabel;
//
//    @FXML
//    private TextField saveAsLocation;
//
//    private List<File> documents = new ArrayList<>();
//    private final  List<File> wordFiles = new ArrayList<>();
//    private final List<File> pptFiles = new ArrayList<>();
//    private final List<File> excelFiles = new ArrayList<>();
//
//
//    @FXML
//    private void initialize() throws Exception {
//
////        inheritSize = inheritCheck.isSelected();
//        FileChooser.ExtensionFilter jpegFilter = new FileChooser.ExtensionFilter("JPEG (*.JPG;*.JPEG;*JPE)", "*.JPG", "*.JPEG", "*.JPE");
//        FileChooser.ExtensionFilter epsFilter = new FileChooser.ExtensionFilter("EPS TIFF Preview (*.EPS)", "*.EPS");
//        FileChooser.ExtensionFilter pngFilter = new FileChooser.ExtensionFilter("PMG (*.PNG)", "*.PNG");
//        FileChooser.ExtensionFilter svgFilter = new FileChooser.ExtensionFilter("SVG (*.SVG;*.SVGZ)", "*.SVG", "*.SVGZ");
//        imageChooser.getExtensionFilters().addAll(jpegFilter, epsFilter, pngFilter, svgFilter);
//
//        FileChooser.ExtensionFilter templateFilter = new FileChooser.ExtensionFilter("PPT Template (*.potx)", "*.potx");
//        templateChooser.getExtensionFilters().addAll(templateFilter);
//
//        FileChooser.ExtensionFilter officeFilter = new FileChooser.ExtensionFilter("Microsoft Office Files",
//                "*.docx", "*.doc", "*.pptx", "*.ppt", "*.xlsx", "*.xlsm", "*.xlsb", "*.xltx", "*.xltm", "*.xls", "*.xlt", "*.xml");
//        FileChooser.ExtensionFilter docFilter = new FileChooser.ExtensionFilter("Word Documents (*.docx;*.doc)", "*.docx", "*.doc");
//        FileChooser.ExtensionFilter pptFilter = new FileChooser.ExtensionFilter("PowerPoint Presentations (*.pptx;*.ppt)", "*.pptx", "*.ppt");
//        FileChooser.ExtensionFilter excelFilter = new FileChooser.ExtensionFilter("Excel Files (*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.xlt;*.xml",
//                "*.xlsx", "*.xlsm", "*.xlsb", "*.xltx", "*.xltm", "*.xls", "*.xlt", "*.xml");
//        docChooser.getExtensionFilters().addAll(officeFilter, docFilter, pptFilter, excelFilter);
//
//    }
//
//    @FXML
//    public void processUpdate(MouseEvent mouseEvent) throws Exception {
//
//        updateWordFiles(wordFiles);
//        updatePowerPointFiles(pptFiles);
//        updateExcelFiles(excelFiles);
////        saveDirectory = new File(saveAsLocation.getText());
//
//        updateCompleteLabel.setVisible(true);
//        resetGrid();
//
//    }
//
//    private void resetGrid() {
//        // disable grids
//        leftGrid.setDisable(true);
//        rightGrid.setDisable(true);
//        updateLogoGrid.setDisable(true);
//        updateLinksGrid.setDisable(true);
//
//        //deselect buttons
//        updateLinksBtn.setSelected(false);
//        updateLogoBtn.setSelected(false);
//        updateTemplateBtn.setSelected(false);
//        selectTemplateFile.setDisable(true);
//        updateBtn.setDisable(true);
//
//        // disable buttons/labels
//        updateTemplateBtn.setDisable(true);
//        updateLinksBtn.setDisable(true);
////        selectDirectory.setDisable(true);
//        loadedLabel.setVisible(false);
//
//        // clear replacement list
//        if(!replacementList.getItems().isEmpty()) {
//            replacementList.getItems().clear();
//        }
//        oldText.clear();
//        newText.clear();
//    }
//
//    private void updateWordFiles(List<File> files) {
//        if (files.isEmpty())
//            return;
//
//        System.out.println("Updating Word Files...");
//        List<Document> documents = new ArrayList<>();
//        for(File file : files) {
//            try {
//                documents.add(new Document(file.getAbsolutePath()));
//            } catch (Exception e) {
//                System.out.println("\tLoad Document Exception: " + e.getMessage());
//            }
//        }
////        files.forEach(file -> {
////            try {
////                documents.add(new Document(file.getAbsolutePath()));
////            } catch (Exception e) {
////                System.out.println("\tLoad Document Exception: " + e.getMessage());
////            }
////        });
//
//        for (int i = 0; i < documents.size(); i++) {
//
//            Document document = documents.get(i);
//
//            if (updateLinksBtn.isSelected()) {
//                System.out.println("\tUpdating Hyperlinks");
//                for (Section section : (Iterable<Section>) document.getSections()) {
//                    for (Paragraph paragraph : (Iterable<Paragraph>) section.getParagraphs()) {
//                        for (DocumentObject object : (Iterable<DocumentObject>) paragraph.getChildObjects()) {
//                            if (object.getDocumentObjectType().equals(DocumentObjectType.Field)) { // check if object is a field
//                                Field field = (Field) object;
//
//                                // check if field is a hyperlink and matches the address to be replaced
//                                if (field.getType().equals(FieldType.Field_Hyperlink) &&
//                                        field.getCode().contains(oldHyperlinkField.getCharacters())) {
//                                    field.setCode(getNewHyperlink(field.getCode()));
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//
//            if (updateLogoBtn.isSelected()) {
//                System.out.println("\tUpdating Logos");
//                for (Section section : (Iterable<Section>) document.getSections()) {
//                    //check headers/footers first
//                    HeaderFooter header = section.getHeadersFooters().getHeader();
//                    HeaderFooter footer = section.getHeadersFooters().getFooter();
//
//                    updateLogosWord(header.getParagraphs());
//                    updateLogosWord(footer.getParagraphs());
//                    updateLogosWord(section.getParagraphs());
//                }
//            }
//
//
//            if(!replacementList.getItems().isEmpty()) {
//                for(Label label : replacementList.getItems()) {
//                    String[] replace = label.getText().split(" → ");
//                    document.replace(Pattern.compile(replace[0]), replace[1]);
//                }
//            }
//
////            String name = saveDirectory.getAbsolutePath() + "/" + files.get(i).getName().replace(".docx", "-test.docx");
//            String name = saveAsLocation.getText() + "/" + files.get(i).getName().replace(".docx", "-test.docx");
//            document.saveToFile(name, documents.get(i).getDetectedFormatType());
//        }
//    }
//
//    private void updatePowerPointFiles(List<File> files) throws Exception {
//        if (files.isEmpty())
//            return;
//        System.out.println("Updating PowerPoint Files...");
//        List<Presentation> presentations = new ArrayList<>();
//        for(File file : files) {
//            try {
//                Presentation presentation = new Presentation();
//                presentation.loadFromFile(file.getAbsolutePath());
//                presentations.add(presentation);
//            } catch (Exception e) {
//                System.out.println("Presentation Load Error: " + e.getMessage());
//            }
//        }
////        files.forEach(file -> {
////            try {
////                Presentation presentation = new Presentation();
////                presentation.loadFromFile(file.getAbsolutePath());
////                presentations.add(presentation);
////            } catch (Exception e) {
////                System.out.println("Presentation Load Error: " + e.getMessage());
////            }
////        });
//
//        for (int i = 0; i < presentations.size(); i++) {
//
//            Presentation presentation = presentations.get(i);
//
//            // update PPT Hyperlinks & words
//            if (updateLinksBtn.isSelected()) {
//                System.out.println("\tUpdating Hyperlinks");
//                for (Object slide : presentation.getSlides()) {
//                    if (slide instanceof ISlide) {
//                        for (Object shape : ((ISlide) slide).getShapes()) {
//                            if (shape instanceof IShape) {
//                                if (((IShape) shape).getClick() != null) { //replace hyperlink embedded in image/shape
//                                    if (((IShape) shape).getClick().getAddress() != null) {
//                                        String hyperlink = ((IShape) shape).getClick().getAddress();
//                                        if (hyperlink.contains(oldHyperlinkField.getCharacters())) {
//                                            ((IShape) shape).getClick().setAddress(getNewHyperlink(hyperlink));
//                                        }
//
//                                    }
//                                }
//                                else { // replace hyperlink text and words
//                                    if(shape instanceof IAutoShape) {
//                                        for(Object paragraph : ((IAutoShape) shape).getTextFrame().getParagraphs()) {
//                                            ParagraphEx pe = (ParagraphEx)paragraph;
//                                            for(Object textPortion : pe.getTextRanges()) {
//                                                PortionEx te = (PortionEx) textPortion;
//
//                                                if(te.getClickAction() != null) { // change hyperlink address
//                                                    System.out.println("hyperlink name: "+((IShape) shape).getName());
//                                                    String hyperlink = te.getClickAction().getAddress();
//                                                    if(hyperlink != null && hyperlink.contains(oldHyperlinkField.getCharacters())) {
//                                                        te.getClickAction().setAddress(getNewHyperlink(hyperlink));
//                                                    }
//                                                }
//
//                                                if(!replacementList.getItems().isEmpty()) { // replace words
//                                                    for(Label label : replacementList.getItems()) {
//                                                        String[] replace = label.getText().split(" → ");
//                                                        if(te.getText().contains(replace[0])) {
//                                                            te.setText(te.getText().replace(replace[0], replace[1]));
//                                                        }
//                                                    }
//                                                }
//                                            }
//                                        }
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//
//            // update PPT Templates
//            if (updateTemplateBtn.isSelected()) {
//                System.out.println("\tUpdating Templates...");
//                Presentation template = new Presentation();
//                template.loadFromFile(templateFile.getAbsolutePath());
//                MasterSlideCollection masterSlides = presentation.getMasters();
//                masterSlides.appendSlide(template.getMasters().get(0));
//
//                // set title layouts
//                ISlide firstSlide = presentation.getSlides().get(0);
//
//                firstSlide
//                        .setLayout(masterSlides.get(masterSlides.getCount() - 1)   // get replacement master slide
//                                .getLayouts().get(0));                          // get title layout
//
//                if (firstSlide.getTitle().equals("")) { // title is contained within a textbox
//                    Object shape = firstSlide.getShapes().get(1);
//                    if (shape instanceof IAutoShape) {
//                        IAutoShape titleContainer = (IAutoShape) shape;
//                        firstSlide.setTitle(titleContainer.getTextFrame()
//                                .getParagraphs().get(0).getText());
//
//                        firstSlide.getShapes().removeAt(1); // remove duplicate text
//                    } else {
//                        System.out.println("Title could not be found");
//                    }
//
//                    for (int j = 0; j < firstSlide.getShapes().size(); j++) {
//                        shape = firstSlide.getShapes().get(j);
//                        if (shape instanceof SlidePicture) {
//                            firstSlide.getShapes().removeAt(j);
//                        }
//                    }
//                }
//
//                // set all other slides layout
//                for (int j = 1; j < presentation.getSlides().getCount(); j++) {
//                    ISlide slide = presentation.getSlides().get(j);
//                    List<Float> positions = new ArrayList<>();
//
//                    // store positions for each shape
//                    for (Object shape : slide.getShapes()) {
//                        if (shape instanceof IAutoShape) {
//                            positions.add(((IAutoShape) shape).getLeft());
//                        } else if (shape instanceof IShape) {
//                            positions.add(((IShape) shape).getLeft());
//                        }
//                    }
//                    slide.setLayout(masterSlides.get(masterSlides.getCount() - 1)
//                            .getLayouts().get(3));
//
//                    // re-apply the positions
//                    for (int k = 0; k < positions.size(); k++) {
//                        slide.getShapes().get(k)
//                                .setLeft(positions.get(k));
//                    }
//                }
//            }
//
////            String name = saveDirectory.getAbsolutePath() + "/" + files.get(i).getName().replace(".pptx", "-test.pptx");
//            String name = saveAsLocation.getText() + "/" + files.get(i).getName().replace(".pptx", "-test.pptx");
//            presentation.saveToFile(name, com.spire.presentation.FileFormat.PPTX_2013);
//        }
//    }
//
//    private void updateExcelFiles(List<File> files) {
//        if (files.isEmpty())
//            return;
//
//        System.out.println("Updating Excel Files...");
//        List<Workbook> workbooks = new ArrayList<>();
//        for(File file : files) {
//            try {
//                Workbook wb = new Workbook();
//                wb.loadFromFile(file.getAbsolutePath());
//                workbooks.add(wb);
//            } catch (Exception e) {
//                System.out.println("Error loading Excel Workbook: " + e.getMessage());
//            }
//        }
////        files.forEach(file -> {
////            try {
////                Workbook wb = new Workbook();
////                wb.loadFromFile(file.getAbsolutePath());
////                workbooks.add(wb);
////            } catch (Exception e) {
////                System.out.println("Error loading Excel Workbook: " + e.getMessage());
////            }
////        });
//
//        for (int i = 0; i < workbooks.size(); i++) {
//            Workbook wb = workbooks.get(i);
//
//            for(Object ws : wb.getWorksheets()) {
//                if(ws instanceof Worksheet) {
//
//                    if(updateLinksBtn.isSelected()) { // update hyperlink
//                        System.out.println("\tUpdating Hyperlinks");
//                        for (Object hyperLink : ((Worksheet) ws).getHyperLinks()) {
//                            if (hyperLink instanceof HyperLink) {
//                                if (((HyperLink) hyperLink).getAddress().contains(oldHyperlinkField.getCharacters())) {
//                                    ((HyperLink) hyperLink).setAddress(getNewHyperlink(((HyperLink) hyperLink).getAddress()));
//                                }
//                            }
//                        }
//                    }
//
//                    if(updateLogoBtn.isSelected()) {
//                        for (Object picture : ((Worksheet) ws).getPictures()) {
//                            if (picture instanceof ExcelPicture) {
//                                if (((ExcelPicture) picture).getAlternativeText().contains("har")) {
//                                    BufferedImage img;
//                                    try {
//                                        img = ImageIO.read(pictureFile);
//                                        ((ExcelPicture) picture).setPicture(img);
//                                        ((ExcelPicture) picture).setAlternativeText("l3har_logo");
//                                    } catch (IOException e) {
//                                        e.printStackTrace();
//                                    }
//                                }
//                            }
//                        }
//                    }
//
//                    for(Label label : replacementList.getItems()) { // replace words
//                        String[] replace = label.getText().split(" → ");
//                        for(CellRange range : ((Worksheet) ws).findAllString(replace[0], true, true)) {
//                            range.setText(range.getText().replace(replace[0], replace[1]));
//                        }
//                    }
//
//                }
//            }
//
////            String name = saveDirectory.getAbsolutePath() + "/" + files.get(i).getName().replace(".xlsx", "-test.xlsx");
//            String name = saveAsLocation.getText() + "/" + files.get(i).getName().replace(".xlsx", "-test.xlsx");
//            wb.saveToFile(name);
//        }
//    }
//
//    private void updateLogosWord(Iterable<Paragraph> paragraphs) {
//        for (Paragraph paragraph : paragraphs) {
//            for (DocumentObject object : (Iterable<DocumentObject>) paragraph.getChildObjects()) {
//                if (object instanceof DocPicture) {
//                    System.out.println("docpicture");
//                    DocPicture picture = (DocPicture) object;
//                    if (picture.getAlternativeText().contains("har")) {
//                        System.out.println("logo picture found");
//                        picture.loadImage(pictureFile.getAbsolutePath());
//                        picture.setAlternativeText("l3har_logo");
//                    }
//                }
//            }
//        }
//    }
//
//    @FXML
//    public void addReplaceWord(MouseEvent mouseEvent) {
//        Label label = new Label();
//        label.setText(oldText.getText() + " → "+ newText.getText());
//        replacementList.getItems().add(label);
//        oldText.clear();
//        newText.clear();
//        checkUpdateButton();
//    }
//
//    @FXML
//    public void removeReplaceWord(MouseEvent mouseEvent) {
//        replacementList.getItems().remove(
//                replacementList.getSelectionModel().getSelectedIndex());
//        checkUpdateButton();
//    }
//
//    private String getNewHyperlink(String original) {
//        return original.replace(oldHyperlinkField.getText(), newHyperlinkField.getText());
//    }
//
//
//    private void checkUpdateButton() {
//        if(!updateLogoBtn.isSelected() && !updateLinksBtn.isSelected() && !updateTemplateBtn.isSelected()) {
//            if(!replacementList.getItems().isEmpty()) {
//                updateBtn.setDisable(false);
////                selectDirectory.setDisable(false);
//            } else {
//                updateBtn.setDisable(true);
////                selectDirectory.setDisable(true);
//            }
//        }
//        else {
//            if(updateLogoBtn.isSelected() && pictureFile != null && updateBtn.isDisabled()) {
//                updateBtn.setDisable(false);
////                selectDirectory.setDisable(false);
//            }
//            if(updateTemplateBtn.isSelected() && templateFile != null) {
//                updateBtn.setDisable(false);
////                selectDirectory.setDisable(false);
//            }
//            if(updateLinksBtn.isSelected()) {
//                if(!newHyperlinkField.getText().equals("") && !oldHyperlinkField.getText().equals("")) {
//                    updateBtn.setDisable(false);
////                    selectDirectory.setDisable(false);
//                }
//                else {
//                    updateBtn.setDisable(true);
////                    selectDirectory.setDisable(true);
//                }
//            }
//        }
//    }
//
//    @FXML
//    public void oldHyperlinkFieldUpdate(KeyEvent keyEvent) {
//        checkUpdateButton();
//    }
//
//    @FXML
//    public void newHyperlinkFieldUpdate(KeyEvent keyEvent) {
//        checkUpdateButton();
//    }
//
//    @FXML
//    public void toggleTemplateBtn(MouseEvent mouseEvent) {
//        selectTemplateFile.setDisable(!updateTemplateBtn.isSelected());
//        checkUpdateButton();
//    }
//
//    @FXML
//    public void toggleLinksUpdates(MouseEvent mouseEvent) {
//        updateLinksGrid.setDisable(!updateLinksBtn.isSelected());
//        checkUpdateButton();
//    }
//
//    @FXML
//    public void toggleLogoUpdates(MouseEvent mouseEvent) {
//        updateLogoGrid.setDisable(!updateLogoBtn.isSelected());
//        checkUpdateButton();
//    }
//
//
//    @FXML
//    public void loadDocuments(MouseEvent mouseEvent) {
//        documents = docChooser.showOpenMultipleDialog(updateLinksGrid.getScene().getWindow());
//        if (!documents.isEmpty()) {
//            leftGrid.setDisable(false);
//            updateLinksBtn.setDisable(false);
//            rightGrid.setDisable(false);
//            updateLogoBtn.setDisable(false);
//            loadedLabel.setVisible(true);
//            if(updateCompleteLabel.isVisible()) {
//                updateCompleteLabel.setVisible(false);
//            }
//
//            for(File document : documents) {
//                if(document.getName().contains(".doc")) {
//                    wordFiles.add(document);
//                }
//
//                else if(document.getName().contains(".ppt")) {
//                    pptFiles.add(document);
//                }
//
//                else if(document.getName().contains(".xl")) {
//                    excelFiles.add(document);
//                }
//            }
////            wordFiles = documents.stream()
////                    .filter((s) -> s.getName().contains(".doc"))
////                    .collect(Collectors.toList());
//
////            pptFiles = documents.stream()
////                    .filter((s) -> s.getName().contains(".ppt"))
////                    .collect(Collectors.toList());
////
////            excelFiles = documents.stream()
////                    .filter((s) -> s.getName().contains(".xl"))
////                    .collect(Collectors.toList());
//
//            if (!pptFiles.isEmpty()) {
//                updateTemplateBtn.setDisable(false);
//            }
//        }
//    }
//
////    @FXML
////    public void selectDirectory(MouseEvent mouseEvent) {
//////        saveDirectory = directoryChooser.showDialog(updateLinksGrid.getScene().getWindow());
////        System.out.println("Directory Selected: "+saveDirectory.getAbsolutePath());
////    }
//
//    @FXML
//    public void loadTemplate(MouseEvent mouseEvent) {
//        templateFile = templateChooser.showOpenDialog(updateLinksGrid.getScene().getWindow());
//        checkUpdateButton();
//    }
//
//    @FXML
//    public void loadPicture(MouseEvent mouseEvent) {
//        pictureFile = imageChooser.showOpenDialog(updateLogoGrid.getScene().getWindow());
//        Image previewImage = new Image("file:" + pictureFile.getAbsolutePath());
//        imageView.setImage(previewImage);
//        checkUpdateButton();
//
//    }
//
//}