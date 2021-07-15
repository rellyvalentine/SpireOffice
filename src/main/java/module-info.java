module SpireOffice {
    requires javafx.fxml;
    requires javafx.controls;
    requires java.datatransfer;
    requires java.desktop;
    requires jakarta.activation;
    requires java.xml.bind;
    requires java.sql;

    requires spire.doc.free;

    opens sample;
}