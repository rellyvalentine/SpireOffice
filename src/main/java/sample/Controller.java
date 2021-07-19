package sample;

import com.spire.doc.*;
import com.spire.doc.FieldType;
import com.spire.doc.Section;
import com.spire.doc.documents.DocumentObjectType;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.Field;
import com.spire.presentation.*;
import com.spire.presentation.collections.MasterSlideCollection;
import com.spire.presentation.collections.SlideCollection;
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
    private final FileChooser templateChooser = new FileChooser();
    private final FileChooser docChooser = new FileChooser();

    private File pictureFile;
    private File templateFile;

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
    private ToggleButton updateTemplateBtn;

    @FXML
    private Button selectTemplateFile;

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

    private List<File> wordFiles;
    private List<File> pptFiles;
    private List<File> excelFiles;

    @FXML
    private void initialize() throws Exception {

        inheritSize = inheritCheck.isSelected();
        FileChooser.ExtensionFilter jpegFilter = new FileChooser.ExtensionFilter("JPEG (*.JPG;*.JPEG;*JPE)", "*.JPG", "*.JPEG", "*.JPE");
        FileChooser.ExtensionFilter epsFilter = new FileChooser.ExtensionFilter("EPS TIFF Preview (*.EPS)", "*.EPS");
        FileChooser.ExtensionFilter pngFilter = new FileChooser.ExtensionFilter("PMG (*.PNG)", "*.PNG");
        FileChooser.ExtensionFilter svgFilter = new FileChooser.ExtensionFilter("SVG (*.SVG;*.SVGZ)", "*.SVG", "*.SVGZ");
        imageChooser.getExtensionFilters().addAll(jpegFilter, epsFilter, pngFilter, svgFilter);

        FileChooser.ExtensionFilter templateFilter = new FileChooser.ExtensionFilter("PPT Template (*.potx)", "*.potx");
        templateChooser.getExtensionFilters().addAll(templateFilter);

        FileChooser.ExtensionFilter docFilter = new FileChooser.ExtensionFilter("Word Documents (*.docx;*.doc)", "*.docx", "*.doc");
        FileChooser.ExtensionFilter pptFilter = new FileChooser.ExtensionFilter("PowerPoint Presentations (*.pptx;*.ppt)", "*.pptx", "*.ppt");
        FileChooser.ExtensionFilter excelFilter = new FileChooser.ExtensionFilter("Excel Files (*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.xlt;*.xml",
                "*.xlsx", "*.xlsm", "*.xlsb", "*.xltx", "*.xltm", "*.xls", "*.xlt", "*.xml");
        docChooser.getExtensionFilters().addAll(docFilter, pptFilter, excelFilter);

    }

    @FXML
    private void processUpdate() throws Exception {

        updateWordFiles(wordFiles);
        updatePowerPointFiles(pptFiles);
        updateExcelFiles(excelFiles);

    }

    private void updateWordFiles(List<File> files) {
        if(files.isEmpty())
            return;
        List<Document> documents = new ArrayList<>();
        files.forEach(file -> {
            try {
                documents.add(new Document(file.getAbsolutePath()));
            } catch (Exception e) {
                System.out.println("Load Document Exception: "+e.getMessage());
            }
        });


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

    private void updatePowerPointFiles(List<File> files) throws Exception {
        if(files.isEmpty())
            return;
        List<Presentation> presentations = new ArrayList<>();
        files.forEach(file -> {
            try {
                Presentation presentation = new Presentation();
                presentation.loadFromFile(file.getAbsolutePath());
                presentations.add(presentation);
            } catch (Exception e) {
                System.out.println("Presentation Load Error: "+e.getMessage());
            }
        });

        if(updateLinksBtn.isSelected()) {

        }

        if(updateLogoBtn.isSelected()) {

        }

        if(updateTemplateBtn.isSelected()) {
            Presentation template = new Presentation();
            template.loadFromFile(templateFile.getAbsolutePath());
            System.out.println("Template loaded");

            for(int i = 0; i < presentations.size(); i++) {
                MasterSlideCollection masterSlides = presentations.get(i).getMasters();
                masterSlides.appendSlide(template.getMasters().get(0));

                // set title layouts
                Presentation currentPres = presentations.get(i);
                ISlide firstSlide = currentPres.getSlides().get(0);

               firstSlide
                       .setLayout(masterSlides.get(masterSlides.getCount()-1)   // get replacement master slide
                                .getLayouts().get(0));                          // get title layout

                if(firstSlide.getTitle().equals("")) {                          // title is contained within a textbox
                    Object shape = firstSlide.getShapes().get(1);
                    if(shape instanceof IAutoShape) {
                        IAutoShape titleContainer = (IAutoShape) shape;
                        firstSlide.setTitle(titleContainer.getTextFrame()
                                .getParagraphs().get(0).getText());

                        firstSlide.getShapes().removeAt(1); // remove duplicate text
                    } else {
                        System.out.println("Title could not be found");
                    }

                    for(int j = 0; j < firstSlide.getShapes().size(); j++) {
                        shape = firstSlide.getShapes().get(j);
                        if(shape instanceof SlidePicture) {
                            firstSlide.getShapes().removeAt(j);
                        }
                    }
                }
//
                // set all other slides layout
                for(int j = 1; j < currentPres.getSlides().getCount(); j++) {
                    currentPres.getSlides().get(j)
                            .setLayout(masterSlides.get(masterSlides.getCount()-1)
                                    .getLayouts().get(2));
                }

                String name = files.get(i).getAbsolutePath().replace(".pptx", "-test.pptx");
                currentPres.saveToFile(name, com.spire.presentation.FileFormat.PPTX_2013);
            }
        }
    }

    private void updateExcelFiles(List<File> files) {

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

    @FXML
    private void toggleTemplateBtn() {
        selectTemplateFile.setDisable(!updateTemplateBtn.isSelected());
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

            wordFiles = documents.stream()
                    .filter((s) -> s.getName().contains(".doc"))
                    .collect(Collectors.toList());

            pptFiles = documents.stream()
                    .filter((s) -> s.getName().contains(".ppt"))
                    .collect(Collectors.toList());

            excelFiles = documents.stream()
                    .filter((s) -> s.getName().contains(".xl"))
                    .collect(Collectors.toList());

            if(!pptFiles.isEmpty()) {
                updateTemplateBtn.setDisable(false);
            }
        }
    }

    @FXML
    private void loadTemplate() {
        templateFile = templateChooser.showOpenDialog(updateLinksGrid.getScene().getWindow());
    }

    @FXML
    private void loadPicture() {
        pictureFile = imageChooser.showOpenDialog(updateLogoGrid.getScene().getWindow());
        Image previewImage = new Image("file:"+pictureFile.getAbsolutePath());
        imageView.setImage(previewImage);
    }


}