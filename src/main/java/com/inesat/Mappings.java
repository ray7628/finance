package com.inesat;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.List;

@Data
@Component
@ConfigurationProperties(prefix = "mapping")
public class Mappings {
    private String filename;
    private String[] projects;
    private String incomeSheet;
    private String costSheet;

    private List<ColMapping> incomeCols;
    private List<ColMapping> costCols;
}

@Data
class ColMapping {
    private String from;
    private String to;
    private Integer colIndex;
}
