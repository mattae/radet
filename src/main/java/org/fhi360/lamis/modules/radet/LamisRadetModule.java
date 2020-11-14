package org.fhi360.lamis.modules.radet;

import com.foreach.across.core.AcrossModule;
import com.foreach.across.core.annotations.AcrossDepends;
import com.foreach.across.core.context.configurer.ComponentScanConfigurer;
import com.foreach.across.modules.hibernate.jpa.AcrossHibernateJpaModule;

@AcrossDepends(required = AcrossHibernateJpaModule.NAME)
public class LamisRadetModule extends AcrossModule {
    public static final String NAME = "FHIRadetModule";

    public LamisRadetModule() {
        super();
        addApplicationContextConfigurer(
                new ComponentScanConfigurer(getClass().getPackage().getName() + ".web",
                        getClass().getPackage().getName() + ".mapper", getClass().getPackage().getName() + ".service"));
    }

    @Override
    public String getName() {
        return NAME;
    }
}
