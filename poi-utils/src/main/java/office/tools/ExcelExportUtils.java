package office.tools;

import office.anno.ExcelExportEntity;
import office.beans.TitleBean;
import office.enums.PatternStyle;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;
import java.util.List;


/**
 * excel导出工具
 */
public class ExcelExportUtils {

    public static Workbook exportExcel(Class clazz, List data, PatternStyle patternStyle, String sheetName, String tableName) {
        Workbook workbook = getWorkBook(patternStyle);

        //存放title （按顺序）
        List<TitleBean> titles = Lists.newArrayList();

        Field[] declaredFields = clazz.getDeclaredFields();

        int titleRowNum = 1;
        boolean firstSwitch = true;
        boolean secSwitch = true;

        for (Field field : declaredFields) {
            TitleBean titleBean = new TitleBean();
            ExcelExportEntity exportEntityAnnotation = field.getAnnotation(ExcelExportEntity.class);
            if (exportEntityAnnotation != null) {
                if (!exportEntityAnnotation.isDynamicTitle()) {
                    if (firstSwitch && !StringUtils.isEmpty(exportEntityAnnotation.parentName())) {
                        titleRowNum++;
                        firstSwitch = false;
                    }
                    titleBean.setName(exportEntityAnnotation.name());
                    titleBean.setDynamicTitle(exportEntityAnnotation.isDynamicTitle());
                    titleBean.setOrder(exportEntityAnnotation.order());
                    titleBean.setParentName(exportEntityAnnotation.parentName());
                    titles.add(titleBean);
                } else {
                    //是动态列 根据传来的数据据整理出所有包含的动态列
//                    b = 1;
                    if (!StringUtils.isEmpty(exportEntityAnnotation.parentName())) {
                        titleRowNum++;
                    }

                    //TODO
                }
            }

        }
        System.out.println(titles);

        //行
//        titleRowNum = titleRowNum + a + b;

        Sheet sheet = workbook.createSheet(sheetName);
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(tableName);


        //重新排序后的
        List<TitleBean> sortTitles = Lists.newArrayList();

        if (CollectionUtils.isNotEmpty(titles)) {
            //所有父节点的排序号
//            List<Integer> sortNum = titles.stream().map(TitleBean::getParentOrder).sorted(Integer::compareTo).distinct().collect(Collectors.toList());
//            for (Integer num : sortNum) {
//                List<TitleBean> titleBeanList = titles.stream().filter(e -> e.getParentOrder() == num).sorted(Comparator.comparing(TitleBean::getOrder)).collect(Collectors.toList());
//                sortTitles.addAll(titleBeanList);
//            }

            String bigParentTitle = sortTitles.get(0).getParentName();
            String middleParentTitle = sortTitles.get(0).getName();

            int startColNum = 0;
            int endColNum = 0;


            Row row1 = sheet.createRow(1);
            Row row2 = sheet.createRow(2);
            Row row3 = sheet.createRow(3);

            for (int i = 0; i < sortTitles.size(); i++) {

                TitleBean titleBean = sortTitles.get(i);
                if (!StringUtils.isEmpty(titleBean.getParentName())) {
                    //最后一次循环
                    if (i == sortTitles.size() - 1) {

                        if (titleBean.getParentName().equals(bigParentTitle)) {
                            cell = row1.createCell(startColNum);
                            cell.setCellValue(bigParentTitle);
                            if ((sortTitles.size() - 1 - startColNum) > 1) {
                                sheet.addMergedRegion(new CellRangeAddress(1, titleRowNum - 1, startColNum, sortTitles.size() - 1));
                            }
                        } else {

                            cell = row1.createCell(startColNum);
                            cell.setCellValue(bigParentTitle);
                            if ((sortTitles.size() - 2 - startColNum) > 1) {
                                sheet.addMergedRegion(new CellRangeAddress(1, titleRowNum - 1, startColNum, sortTitles.size() - 2));
                            }

                            cell = row1.createCell(sortTitles.size() - 1);
                            cell.setCellValue(titleBean.getParentName());
                        }

                    } else {
                        if (!titleBean.getParentName().equals(bigParentTitle)) {
                            cell = row1.createCell(startColNum);
                            cell.setCellValue(bigParentTitle);
                            if ((endColNum - startColNum) > 1) {
                                sheet.addMergedRegion(new CellRangeAddress(1, titleRowNum - 1, startColNum, endColNum));
                            }
                            startColNum = endColNum + 1;
                            bigParentTitle = titleBean.getParentName();
                        }
                    }

                    //有父节点 最低行数为2
                    cell = titleRowNum == 2 ? row2.createCell(endColNum) : row3.createCell(endColNum);
                    cell.setCellValue(titleBean.getName());
                    endColNum++;

                } else {

                    cell = (titleRowNum == 2 ? row2.createCell(startColNum) : row3.createCell(startColNum));
                    cell.setCellValue(titleBean.getName());
                    bigParentTitle = (i < sortTitles.size() - 1) ? sortTitles.get(i + 1).getParentName() : "";

                    sheet.addMergedRegion(new CellRangeAddress(1, titleRowNum, startColNum, endColNum));

                    startColNum = endColNum + 1;

                    endColNum++;
                }

            }

        }

        return workbook;
    }

    private static Workbook getWorkBook(PatternStyle patternStyle) {
        if (PatternStyle.XLSX.equals(patternStyle)) {
            return new XSSFWorkbook();
        }
        if (PatternStyle.XLS.equals(patternStyle)) {
            return new HSSFWorkbook();
        }
        throw new RuntimeException("no such this pattern !");
    }

}
