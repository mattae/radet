package org.fhi360.lamis.modules.radet.util;

import org.apache.commons.lang3.StringUtils;
import org.lamisplus.modules.base.config.ContextProvider;
import org.springframework.jdbc.core.JdbcTemplate;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class RegimenIntrospector {
    private final static JdbcTemplate jdbcTemplate = ContextProvider.getBean(JdbcTemplate.class);

    private RegimenIntrospector() {
    }

    public static String resolveRegimen(String regimensys) {
        String query = "SELECT regimen FROM regimen_resolver WHERE regimensys = ?";
        String regimen = jdbcTemplate.query(query, rs -> {
            if (rs.next()) {
                return rs.getString("regimen");
            }
            return "";
        }, regimensys);
        if (StringUtils.equals("Other", regimen) || StringUtils.equals("", regimen)) {
            String regex = ".*(\\(.*\\)).*";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(regimensys);
            while (matcher.matches()) {
                if (matcher.group(1) != null) {
                    regimensys = regimensys.replace(matcher.group(1), "");
                }
                matcher = pattern.matcher(regimensys);
            }
            regimen = regimensys.replaceAll("/", "-").replaceAll("-r", "/r")
                    .replaceAll("\\+", "-");
        }
        return regimen;
    }
}
