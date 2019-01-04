import com.tm.excel.entity.HandleExcelResult;
import com.tm.excel.entity.out.LeadingExcelResponse;
import com.tm.excel.framework.ExcelFactory;
import com.tm.excel.framework.ProduceExcelInputStream;
import entity.DefaultTestExcelBean;
import entity.OrderRecordingRuleExportExcelVo;
import entity.OrderRecordingRuleLeadingInExcelVo;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @Author TroubleMan
 * @Date 2018/5/10 19:01
 * @Description excel导出工具测试 每个对象必须继承BasicExcelExport类或者实现BaseExcel接口
 *              且每个导出的数据封装对象类上必须标注@ExcelSheet，@ExcelHead，@ExcelWorkBook注解 每个注解属性详情请参照注解对应的注释
 *              导出数据封装对象中可以包含非导出的字段以做其他业务处理 即对象导出的属性的值只包括标注了@ExcelColumn注解的属性
 *              工具暂时只支持一个excel文件中一个sheet
 *              且可以自己设置简单的合并单元格，即在@ExcelColumn和@ExcelHead添加rangeAddress长度为4的一维数组 例如rangeAddress
 *              ={0,1,0,4},表示从第1行到第二行以及从第1列到第5列合并为一个单元格
 **/

public class HaixueExcelTest {

    /**
     * @Description 导出数据并设置默认样式(注解在OrderRecordingRuleExportExcelVo上)
     *              默认样式只需要继承BaseDefaultStyleExcel，不需要样式则可以直接继承BaseSimpleExcel
     *              如果要自己定义样式则分别注解@ExcelHeadStyl，@ExcelColumnStyle，@ExcelTextStyle
     *              或者注解ExcelDefaultStyle自定义统一样式
     * @Author TroubleMan
     * @date 2018/5/11 10:02
     * @param
     * @return
     **/

    //@Test
    public void test1() throws Exception {

        // 传入List<T>和Class<T>返回数据流对象
        InputStream inputStream = ExcelFactory.produceExcelOfInputStream(this.simulationSelectData1(),
                OrderRecordingRuleLeadingInExcelVo.class);

        // 模拟将文件写入test目录下
        File file = new File("../excel-handle/src/test/" + "导出测试默认样式数据" + ".xls");
        OutputStream os = new FileOutputStream(file);
        int bytesRead;
        byte[] buffer = new byte[8192];
        while ((bytesRead = inputStream.read(buffer, 0, 8192)) != -1) {
            os.write(buffer, 0, bytesRead);
        }
        os.close();
        inputStream.close();
    }


    /**
     * @Author zhangyi
     * @Description: 导出数据并不设置样式(注解在OrderRecordingRuleExportExcelVo上)
     * @Date 2018/9/27
     * @Param []
     * @return void
     **/

    // @Test
    public void test2() throws Exception {
        // 传入List<T>和Class<T>返回数据流对象
        InputStream inputStream = ExcelFactory.produceExcelOfInputStream(this.simulationSelectData2(),
                OrderRecordingRuleExportExcelVo.class);

        // 模拟将文件写入test目录下
        File file = new File("../excel-handle/src/test/" + "导出测试无样式数据" + ".xls");
        OutputStream os = new FileOutputStream(file);
        int bytesRead;
        byte[] buffer = new byte[8192];
        while ((bytesRead = inputStream.read(buffer, 0, 8192)) != -1) {
            os.write(buffer, 0, bytesRead);
        }
        os.close();
        inputStream.close();
    }


    /**
     * @Description 导入数据，要求对象中@ExcelColumn注解的name的值与excel表中的列名值一一对应相同， 且如果头标题存在，则添加属性hasHeaderTitle =
     *              true， 且占位占不止一行，则@ExcelSheet中添加属性：headerTitleHeight = 标题所占单元格长度
     *              在业务支持的情况下，导入和导出可以用同一个bean
     * @Author TroubleMan
     * @date 2018/5/11 10:50
     * @param
     * @return
     **/

    // @Test
    public void test3() throws Exception {
        // 模拟导出excel文件
        File file = new File("../excel-handle/src/test/" + "导入测试数据" + ".xls");

        // 传入输入流InputStream和Class<T>对象，返回一个LeadingExcelResponse<T>对象
        LeadingExcelResponse<OrderRecordingRuleLeadingInExcelVo> leadingExcelResponse = ExcelFactory
                .writeExcelOfInputStream(new FileInputStream(file), OrderRecordingRuleLeadingInExcelVo.class);

        // 返回数据包括标题，列名集合以及实际有效业务数据集合
        List<OrderRecordingRuleLeadingInExcelVo> list = leadingExcelResponse.getColumnDataValues();
        for (OrderRecordingRuleLeadingInExcelVo orderRecordingRuleLeadingInExcelVo : list) {
            System.out.println(orderRecordingRuleLeadingInExcelVo);
        }
    }

    /**
     * @Description 获取导出的excel的表sheet结构对象 包括标题头部，当前sheet本身，每个单元格和每个文字属性封装对象，
     *              各自都含有对应的style样式对象，以便于由于业务需要对注解中已设置好的样式进行二次修改和覆用 仅供特殊业务需要
     * @Author TroubleMan
     * @date 2018/5/11 10:50
     * @param
     * @return
     **/
    // @Test
    public void test4() throws Exception {

        HandleExcelResult handleExcelResult =
                ProduceExcelInputStream.initExcelArchitecture(OrderRecordingRuleExportExcelVo.class);

        System.out.println(handleExcelResult);

    }


    /**
     * @Description 模拟数据库查询
     * @Author TroubleMan
     * @date 2018/5/10 19:05
     * @param
     * @return
     **/
    private List<OrderRecordingRuleLeadingInExcelVo> simulationSelectData1() {
        // 模拟数据库查询
        List<OrderRecordingRuleLeadingInExcelVo> list = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            OrderRecordingRuleLeadingInExcelVo orderRecordingRuleLeadingInExcelVo =
                    new OrderRecordingRuleLeadingInExcelVo();
            orderRecordingRuleLeadingInExcelVo.setCompanyName("222222222222222222222");
            orderRecordingRuleLeadingInExcelVo.setNetPayMoney(new BigDecimal(333333333));
            orderRecordingRuleLeadingInExcelVo.setOnlinePayTypeName("444444444444");
            orderRecordingRuleLeadingInExcelVo.setOrderNo("666666666666666666666");
            orderRecordingRuleLeadingInExcelVo.setTradeNo("77777777777777777777");
            list.add(orderRecordingRuleLeadingInExcelVo);
        }
        return list;
    }

    /**
     * @Description 模拟数据库产生数据
     * @Author TroubleMan
     * @date 2018/5/11 10:05
     * @param
     * @return
     **/

    private List<OrderRecordingRuleExportExcelVo> simulationSelectData2() {
        // 模拟数据库查询
        List<OrderRecordingRuleExportExcelVo> list = new ArrayList<>();
        for (int i = 0; i < 200; i++) {
            OrderRecordingRuleExportExcelVo orderRecordingRuleExportExcelVo =
                    new OrderRecordingRuleExportExcelVo();
            orderRecordingRuleExportExcelVo.setAccountNumber("7");
            orderRecordingRuleExportExcelVo.setCompanyName("6");
            orderRecordingRuleExportExcelVo.setNetPayMoney(new BigDecimal(5));
            orderRecordingRuleExportExcelVo.setOnlinePayTypeName("4");
            orderRecordingRuleExportExcelVo.setOperationName("3");
            orderRecordingRuleExportExcelVo.setOrderNo("2");
            orderRecordingRuleExportExcelVo.setTradeNo("1");
            /**
             * 对于时间的处理可以直接设置Date类型变量， 并在对应属性名上面添加属性dateFormat，例如dateFormat = "yyyy-MM-dd HH:mm:ss"
             */
            orderRecordingRuleExportExcelVo.setPayTime(new Date());
            list.add(orderRecordingRuleExportExcelVo);
        }
        return list;
    }

    // @Test
    public void test5() {
        List<DefaultTestExcelBean> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            DefaultTestExcelBean defaultTestExcelBean = new DefaultTestExcelBean();
            defaultTestExcelBean.setId(new Long(i));
            defaultTestExcelBean.setName(i + "");
            list.add(defaultTestExcelBean);
        }
        try {
            InputStream inputStream =
                    ExcelFactory.produceExcelOfInputStream(list, DefaultTestExcelBean.class);
            // 模拟将文件写入test目录下
            File file = new File("../excel-handle/src/test/" + "导出测试默认注解或者继承不同父类数据" + ".xls");
            OutputStream os = new FileOutputStream(file);
            int bytesRead;
            byte[] buffer = new byte[8192];
            while ((bytesRead = inputStream.read(buffer, 0, 8192)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
            os.close();
            inputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // @Test
    public void test6() throws Exception {
        // 传入List<T>和Class<T>返回数据流对象
        InputStream inputStream = ExcelFactory.produceExcelOfInputStream(this.simulationSelectData2());
        // 模拟将文件写入test目录下
        File file = new File("../excel-handle/src/test/" + "导出测试无样式数据" + ".xls");
        OutputStream os = new FileOutputStream(file);
        int bytesRead;
        byte[] buffer = new byte[8192];
        while ((bytesRead = inputStream.read(buffer, 0, 8192)) != -1) {
            os.write(buffer, 0, bytesRead);
        }
        os.close();
        inputStream.close();
    }

}
