package office.beans;

import lombok.Data;
import org.apache.commons.compress.utils.Lists;

import java.util.List;

@Data
public class TitleBean {

    /**
     * 表头名称
     *
     * @return
     */
    private String name;

    /**
     * 对应顺序
     *
     * @return
     */
    private int order;

    /**
     * 是否动态表头
     *
     * @return
     */
    private boolean isDynamicTitle;

    /**
     * 父级title
     *
     * @return
     */
    private String parentName;

    /**
     * 子集
     */
    private List<TitleBean> childrenTitleBean = Lists.newArrayList();
}
