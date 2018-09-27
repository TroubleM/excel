package entity;

import com.haixue.excel.annotation.ExcelBean;
import com.haixue.excel.annotation.ExcelColumn;
import com.haixue.excel.base.BaseDefaultStyleExcel;

@ExcelBean(sheetName = "默认注解ExcelBean测试", headName = "默认注解ExcelBean测试")
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
