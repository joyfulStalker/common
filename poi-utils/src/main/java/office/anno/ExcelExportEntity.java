package office.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelExportEntity {
    /**
     * 表头名称
     *
     * @return
     */
    String name() default "";

    /**
     * 对应顺序
     *
     * @return
     */
    int order() default 0;

    /**
     * 是否动态表头
     *
     * @return
     */
    boolean isDynamicTitle() default false;

    /**
     * 父级title
     *
     * @return
     */
    String parentName() default "";

}