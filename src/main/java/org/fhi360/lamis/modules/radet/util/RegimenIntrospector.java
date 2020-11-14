package org.fhi360.lamis.modules.radet.util;

import org.lamisplus.modules.base.config.ContextProvider;
import org.springframework.jdbc.core.JdbcTemplate;

public class RegimenIntrospector {
    private final static JdbcTemplate jdbcTemplate = ContextProvider.getBean(JdbcTemplate.class);

    private RegimenIntrospector() {
    }

    public static String resolveRegimen(String regimensys) {
        String query = "SELECT regimen FROM regimen_resolver WHERE regimensys = ?";
        return jdbcTemplate.query(query, rs -> {
            if (rs.next()) {
                return rs.getString("regimen");
            }
            return "";
        }, regimensys);
    }
}
