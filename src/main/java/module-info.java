module SpireOffice {
    requires javafx.fxml;
    requires javafx.controls;
    requires java.datatransfer;
    requires java.desktop;
    requires jakarta.activation;
    requires java.xml.bind;
    requires java.sql;

    requires spire.office.free;
//    requires spire.doc.free;
//    requires spire.presentation.free;

    opens sample;
}