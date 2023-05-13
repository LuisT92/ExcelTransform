module com.example.exceltransform {
    requires javafx.controls;
    requires javafx.fxml;

    requires com.dlsc.formsfx;
    requires net.synedra.validatorfx;
    requires org.kordamp.bootstrapfx.core;
    requires java.desktop;
    requires poi.ooxml;
    requires poi;

    opens com.example.exceltransform to javafx.fxml;
    exports com.example.exceltransform;
}