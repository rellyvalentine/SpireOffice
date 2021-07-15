package sample;

import com.spire.doc.*;
import com.spire.doc.documents.DocumentObjectType;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.Field;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.GridPane;
import javafx.stage.FileChooser;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class Controller {

    private final FileChooser imageChooser = new FileChooser();
    private final FileChooser docChooser = new FileChooser();

    private File pictureFile;

    @FXML
    private GridPane leftGrid;

    @FXML
    private GridPane rightGrid;

    @FXML
    private Label loadedLabel;

    @FXML
    private ToggleButton updateLinksBtn;

    @FXML
    private ToggleButton updateLogoBtn;

    @FXML
    private GridPane updateLinksGrid;

    @FXML
    private GridPane updateLogoGrid;

    @FXML
    private TextField oldHyperlink;

    @FXML
    private TextField newHyperlink;

    @FXML
    private CheckBox inheritCheck;

    @FXML
    private Label widthLabel;

    @FXML
    private Spinner<Integer> widthSpinner;

    @FXML
    private Label heightLabel;

    @FXML
    private Spinner<Integer> heightSpinner;

    @FXML
    private ImageView imageView;

    private List<File> documents;
    private boolean inheritSize;

    @FXML
    private void initialize() throws Exception {

        inheritSize = inheritCheck.isSelected();
        FileChooser.ExtensionFilter jpegFilter = new FileChooser.ExtensionFilter("JPEG (*.JPG;*.JPEG;*JPE)", "*.JPG", "*.JPEG", "*.JPE");
        FileChooser.ExtensionFilter epsFilter = new FileChooser.ExtensionFilter("EPS TIFF Preview (*.EPS)", "*.EPS");
        FileChooser.ExtensionFilter pngFilter = new FileChooser.ExtensionFilter("PMG (*.PNG)", "*.PNG");
        FileChooser.ExtensionFilter svgFilter = new FileChooser.ExtensionFilter("SVG (*.SVG;*.SVGZ)", "*.SVG", "*.SVGZ");
        imageChooser.getExtensionFilters().addAll(jpegFilter, epsFilter, pngFilter, svgFilter);

        FileChooser.ExtensionFilter docFilter = new FileChooser.ExtensionFilter("Word Documents (*.docx;*.doc)", "*.docx", "*.doc");
        FileChooser.ExtensionFilter pptFilter = new FileChooser.ExtensionFilter("PowerPoint Presentations (*.pptx;*.ppt)", "*.pptx", "*.ppt");
        FileChooser.ExtensionFilter excelFilter = new FileChooser.ExtensionFilter("Excel Files (*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.xlt;*.xml",
                "*.xlsx", "*.xlsm", "*.xlsb", "*.xltx", "*.xltm", "*.xls", "*.xlt", "*.xml");
        docChooser.getExtensionFilters().addAll(docFilter, pptFilter, excelFilter);

    }

    @FXML
    private void processUpdate() throws Exception {
        List<File> wordFiles;
        List<File> pptFiles;
        List<File> excelFiles;

        wordFiles = documents.stream()
                .filter((s) -> s.getName().contains(".doc"))
                .collect(Collectors.toList());

        pptFiles = documents.stream()
                .filter((s) -> s.getName().contains(".ppt"))
                .collect(Collectors.toList());

        excelFiles = documents.stream()
                .filter((s) -> s.getName().contains(".xl"))
                .collect(Collectors.toList());

        updateWordFiles(wordFiles);
        updatePowerPointFiles(pptFiles);
        updateExcelFiles(excelFiles);

    }

    private void updateWordFiles(List<File> files) throws Exception {
        if(files.isEmpty())
            return;
        List<Document> documents = new ArrayList<>();
        files.forEach((file -> {
            try {
                documents.add(new Document(file.getAbsolutePath()));
            } catch (Exception e) {
                System.out.println("Load Document Exception: "+e.getMessage());
            }
        }));


        if(updateLinksBtn.isSelected()) {
            for (int i = 0; i < documents.size(); i++) {

                for (Section section : (Iterable<Section>) documents.get(i).getSections()) {
                    for(Paragraph paragraph : (Iterable<Paragraph>) section.getParagraphs()) {
                        for(DocumentObject object : (Iterable<DocumentObject>) paragraph.getChildObjects()) {
                            if(object.getDocumentObjectType().equals(DocumentObjectType.Field)) { // check if object is a field
                                Field field = (Field) object;

                                // check if field is a hyperlink and matches the address to be replaced
                                if(field.getType().equals(FieldType.Field_Hyperlink) &&
                                    field.getCode().contains(oldHyperlink.getCharacters())) {
                                    System.out.print(field.getCode());
                                    field.setCode(getNewHyperlink(field.getCode()));
                                    System.out.println(" --> "+field.getCode());
                                }
                            }
                        }
                    }
                }
                documents.get(i).saveToFile(files.get(i).getAbsolutePath(), documents.get(i).getDetectedFormatType());
            }
        }

        if(updateLogoBtn.isSelected()) {
            for(int i = 0; i < documents.size(); i++) {
                for(Section section : (Iterable<Section>) documents.get(i).getSections()) {

                    //check headers/footers first
                    HeaderFooter header = section.getHeadersFooters().getHeader();
                    HeaderFooter footer = section.getHeadersFooters().getFooter();

                    updateLogo(header.getParagraphs());
                    updateLogo(footer.getParagraphs());
                    updateLogo(section.getParagraphs());
                }
                documents.get(i).saveToFile(files.get(i).getAbsolutePath());
            }
        }
    }

    private void updateLogo(Iterable<Paragraph> paragraphs) {
        for(Paragraph paragraph : paragraphs) {
            for(DocumentObject object : (Iterable<DocumentObject>) paragraph.getChildObjects()) {
                if(object instanceof DocPicture) {
                    System.out.println("docpicture");
                    DocPicture picture = (DocPicture) object;
                    if(picture.getAlternativeText().equals("har2_col") || picture.getAlternativeText().equals("l3har_logo")) {
                        System.out.println("logo picture found");
                        picture.loadImage(pictureFile.getAbsolutePath());
                        picture.setAlternativeText("l3har_logo");
                    }
                }
            }
        }
    }

    private String getNewHyperlink(String original) {
        return original.replace(oldHyperlink.getText(), newHyperlink.getText());
    }

    private void updatePowerPointFiles(List<File> files) {

    }

    private void updateExcelFiles(List<File> files) {

    }

    @FXML
    private void toggleLinksUpdates() {
        updateLinksGrid.setDisable(!updateLinksBtn.isSelected());
    }

    @FXML
    private void toggleLogoUpdates() {
        updateLogoGrid.setDisable(!updateLogoBtn.isSelected());
    }

    @FXML
    private void toggleInherit() {
        inheritSize = inheritCheck.isSelected();
        widthLabel.setDisable(inheritSize);
        widthSpinner.setDisable(inheritSize);
        heightLabel.setDisable(inheritSize);
        heightSpinner.setDisable(inheritSize);
    }

    @FXML
    private void loadDocuments() {
        documents = docChooser.showOpenMultipleDialog(updateLinksGrid.getScene().getWindow());
        if(!documents.isEmpty()) {
            leftGrid.setDisable(false);
            rightGrid.setDisable(false);
            loadedLabel.setVisible(true);
        }
    }

    @FXML
    private void loadPicture() {
        pictureFile = imageChooser.showOpenDialog(updateLogoGrid.getScene().getWindow());
        Image previewImage = new Image("file:"+pictureFile.getAbsolutePath());
        imageView.setImage(previewImage);
    }


}