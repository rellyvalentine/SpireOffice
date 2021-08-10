package sample.components;

import javax.swing.*;
import java.awt.event.FocusEvent;
import java.awt.event.FocusListener;

public class HintTextField extends JTextField implements FocusListener {

    private final String hint;
    private boolean showingHint;

    public HintTextField(final String hint) {
        super(hint);
        this.hint = hint;
        this.showingHint = true;
        super.addFocusListener(this);
    }

    @Override
    public void focusGained(FocusEvent e) {
        if(this.getText().isEmpty() || this.getText().equals(hint)) {
            super.setText("");
        }
        showingHint = false;
    }

    public void focusLost(FocusEvent e) {
        if(this.getText().isEmpty()) {
            super.setText(hint);
            this.showingHint = false;
        }
    }

    @Override
    public String getText() {
        return showingHint ? "" : super.getText();
    }

}
