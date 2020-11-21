package office.enums;

import lombok.Getter;

@Getter
public enum PatternStyle {

    XLSX("xlsx"),
    XLS("xls");

    private String suffix;

    PatternStyle(String suffix) {
        this.suffix = suffix;
    }
}
