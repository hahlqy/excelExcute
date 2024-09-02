package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

public class Main {

    private static final int MAX_XLS = 65536;

    private static final int MAX_XLSX = 1048576;

    public static void main(String[] args) throws IOException {

// 指定要合并的 Excel 文件所在的文件夹路径
        String folderPath = "file/";
        List<Path> files=new ArrayList<>();
        Files.walk(Paths.get(folderPath)).filter(Files::isRegularFile)
                .forEach(files::add);
        files.sort(Comparator.comparing(Path::getFileName));
        System.out.println(files);
        int startRow = 2;
        Map<String,MerSpecialFeeVo> result = new HashMap<>();
        for (Path file:files){
            if(file.toString().endsWith(".xls")||file.toString().endsWith(".xlsx")) {
                try (FileInputStream inputStream = new FileInputStream(file.toFile())) {
                    Workbook inputWorkbook = WorkbookFactory.create(inputStream);
                    Sheet inputSheet = inputWorkbook.getSheetAt(0); // 获取第一个工作表
                    int maxRow = file.toString().endsWith(".xls")?MAX_XLS:MAX_XLSX;
                    for(int rowNum =startRow;rowNum<maxRow;rowNum++){
                        Row sourceRow=inputSheet.getRow(rowNum);
                        if(sourceRow == null){
                            continue;
                        }
                        try{
                            MerSpecialFeeVo merSpecialFeeVo = new MerSpecialFeeVo(sourceRow);
                            if(result.containsKey(merSpecialFeeVo.getMerInnerCode())){
                                MerSpecialFeeVo mapVo = result.get(merSpecialFeeVo.getMerInnerCode());
                                mapVo.add(merSpecialFeeVo);
                            }else{
                                result.put(merSpecialFeeVo.getMerInnerCode(),merSpecialFeeVo);
                            }
                        }catch (Exception e){
                            System.out.println("有异常,请检查列 :"+rowNum);
                        }
                    }

                }

            }
        }
        FileInputStream in =new FileInputStream("temp/template.xlsx");
        XSSFWorkbook wk= new XSSFWorkbook(in);
        Sheet sheet=wk.getSheetAt(0);
        Set<String> keySet = result.keySet();
        int recordNum = 1;
        for(String key:keySet){
            Row row = sheet.createRow(recordNum);
            MerSpecialFeeVo merSpecialFeeVo = result.get(key);
            export2Excel(merSpecialFeeVo,row);
            recordNum++;
        }
        try (FileOutputStream outputStream = new FileOutputStream("汇总表.xlsx")) {
            wk.write(outputStream);
            System.out.println("文件合并完成，文件名为：汇总表.xlsx");
        }

    }

    public static void  export2Excel(MerSpecialFeeVo merSpecialFeeVo,Row row){
        Cell merInnerCode = row.createCell(0);
        merInnerCode.setCellValue(merSpecialFeeVo.getMerInnerCode());
        Cell firstUnit = row.createCell(1);
        firstUnit.setCellValue(merSpecialFeeVo.getFirstUnit());
        Cell unitName = row.createCell(2);
        unitName.setCellValue(merSpecialFeeVo.getUnitName());
        Cell merNameCh = row.createCell(3);
        merNameCh.setCellValue(merSpecialFeeVo.getMerNameCh());
        Cell zyhFee = row.createCell(4);
        zyhFee.setCellValue(merSpecialFeeVo.getZyhFee().doubleValue());
        Cell zjmFee = row.createCell(5);
        zjmFee.setCellValue(merSpecialFeeVo.getZjmFee().doubleValue());
        Cell zshFee = row.createCell(6);
        zshFee.setCellValue(merSpecialFeeVo.getZshFee().doubleValue());
        Cell ztdFee = row.createCell(7);
        ztdFee.setCellValue(merSpecialFeeVo.getZtdFee().doubleValue());
        Cell jgSxySy = row.createCell(8);
        jgSxySy.setCellValue(merSpecialFeeVo.getJgSxySy().doubleValue());
        Cell zbs = row.createCell(9);
        zbs.setCellValue(merSpecialFeeVo.getZbs().doubleValue());
        Cell jyZje = row.createCell(10);
        jyZje.setCellValue(merSpecialFeeVo.getJyZje().doubleValue());
        Cell merStatus = row.createCell(11);
        merStatus.setCellValue(merSpecialFeeVo.getMerStatus());

    }

    public static class MerSpecialFeeVo{
        private String merInnerCode;
        private String firstUnit;
        private String unitName;
        private String merNameCh;
        private BigDecimal zyhFee;
        private BigDecimal zjmFee;
        private BigDecimal zshFee;
        private BigDecimal ztdFee;
        private BigDecimal jgSxySy;
        private BigDecimal zbs;
        private BigDecimal jyZje;
        private String merStatus;

        public MerSpecialFeeVo() {
        }

        public void add(MerSpecialFeeVo vo){
            this.zyhFee = this.zyhFee.add(vo.zyhFee);
            this.zshFee = this.zshFee.add(vo.zshFee);
            this.ztdFee = this.ztdFee.add(vo.ztdFee);
            this.jgSxySy = this.jgSxySy.add(vo.jgSxySy);
            this.zbs = this.zbs.add(vo.zbs);
            this.jyZje = this.jyZje.add(vo.jyZje);

        }

        public BigDecimal getZjmFee() {
            return zjmFee;
        }

        public void setZjmFee(BigDecimal zjmFee) {
            this.zjmFee = zjmFee;
        }

        public MerSpecialFeeVo(Row sourceRow) {
            this.merInnerCode = String.valueOf(sourceRow.getCell(0));
            this.firstUnit = String.valueOf(sourceRow.getCell(1));
            this.unitName = String.valueOf(sourceRow.getCell(2));
            this.merNameCh = String.valueOf(sourceRow.getCell(3));
            this.zyhFee = new BigDecimal(String.valueOf(sourceRow.getCell(4)));
            this.zshFee = new BigDecimal(String.valueOf(sourceRow.getCell(5)));
            this.zjmFee = new BigDecimal(String.valueOf(sourceRow.getCell(6)));
            this.ztdFee = new BigDecimal(String.valueOf(sourceRow.getCell(7)));
            this.jgSxySy = new BigDecimal(String.valueOf(sourceRow.getCell(8)));
            this.zbs = new BigDecimal(String.valueOf(sourceRow.getCell(9)));
            this.jyZje = new BigDecimal(String.valueOf(sourceRow.getCell(10)));
            this.merStatus = String.valueOf(sourceRow.getCell(11));
        }

        public String getMerInnerCode() {
            return merInnerCode;
        }

        public void setMerInnerCode(String merInnerCode) {
            this.merInnerCode = merInnerCode;
        }

        public String getFirstUnit() {
            return firstUnit;
        }

        public void setFirstUnit(String firstUnit) {
            this.firstUnit = firstUnit;
        }

        public String getUnitName() {
            return unitName;
        }

        public void setUnitName(String unitName) {
            this.unitName = unitName;
        }

        public String getMerNameCh() {
            return merNameCh;
        }

        public void setMerNameCh(String merNameCh) {
            this.merNameCh = merNameCh;
        }

        public BigDecimal getZyhFee() {
            return zyhFee;
        }

        public void setZyhFee(BigDecimal zyhFee) {
            this.zyhFee = zyhFee;
        }

        public BigDecimal getZshFee() {
            return zshFee;
        }

        public void setZshFee(BigDecimal zshFee) {
            this.zshFee = zshFee;
        }

        public BigDecimal getZtdFee() {
            return ztdFee;
        }

        public void setZtdFee(BigDecimal ztdFee) {
            this.ztdFee = ztdFee;
        }

        public BigDecimal getJgSxySy() {
            return jgSxySy;
        }

        public void setJgSxySy(BigDecimal jgSxySy) {
            this.jgSxySy = jgSxySy;
        }

        public BigDecimal getZbs() {
            return zbs;
        }

        public void setZbs(BigDecimal zbs) {
            this.zbs = zbs;
        }

        public BigDecimal getJyZje() {
            return jyZje;
        }

        public void setJyZje(BigDecimal jyZje) {
            this.jyZje = jyZje;
        }

        public String getMerStatus() {
            return merStatus;
        }

        public void setMerStatus(String merStatus) {
            this.merStatus = merStatus;
        }
    }
}