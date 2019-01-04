package entity;

import com.tm.excel.annotation.ExcelBean;
import com.tm.excel.annotation.ExcelColumn;
import com.tm.excel.base.BaseDefaultStyleExcel;

@ExcelBean(sheetName = "默认注解ExcelBean测试", hasHeaderTitle = true,
        headName = "默认注解ExcelBean测试",headerTitleHeight = 2,headRangeAddress ={0,1,0,4})
public class DefaultTestExcelBean extends BaseDefaultStyleExcel {

    @ExcelColumn(name = "主键")
    private Long id;

    @ExcelColumn(name = "名称")
    private String name;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
