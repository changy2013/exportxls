package org.feiyu.export.xls;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.util.*;

import lombok.extern.slf4j.Slf4j;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.feiyu.export.bean.GoonProduct;
import org.feiyu.export.bean.ProductCompare;
import org.feiyu.export.common.DateUtils;

import java.io.*;

/**
 * @ Author : yangchang@ctrip.com
 * @ Desc ：
 * @ Date : Created in 2017/12/4 18:22
 * @ Modified By ：
 */
@Slf4j
public class ExportXls {

    private static Map<String,String> skuMap = Maps.newHashMap();

    public static void main(String[] args) throws IOException {

        ProductCompare productCompare = new ProductCompare();

        File fileDir = new File("d:/goon/out");

        if (!fileDir.exists()) {
            fileDir.mkdirs();
        }
        Map<String,String> priceMap = readPriceXls();
        Map<String,Integer> transInvMap = readTransInv("ELLEAIR-B2C-TM");
        log.info("生成纸品报表");
        String zhipin = "d:/goon/zhipin.xls";
        List<GoonProduct> zhipinList = readXlsData(zhipin);
        List<GoonProduct> mergeList = mergeProduct(zhipinList,transInvMap);
        addTransProduct(mergeList,transInvMap,"ELLEAIR-B2C-TM");
        mergeList = mergeWareHourse(mergeList);
        replacePrice(mergeList,priceMap);
        Collections.sort(mergeList,productCompare);
        writeFileZhipin(mergeList);

        log.info("生成猫超报表");
        String maochao = "d:/goon/maochao.xls";
        List<GoonProduct> maochaoList = readXlsData(maochao);
        if (maochaoList != null) {
            Map<String,Integer> maoChaotransInvMap = readTransInv("GOON-B2B-DX");
            List<GoonProduct> maochaoMerge = mergeProduct(maochaoList,maoChaotransInvMap);
            List<GoonProduct> maochaoResult = mergeWareHourse(maochaoMerge);
            addTransProduct(maochaoResult,maoChaotransInvMap,"GOON-B2B-DX");
            maochaoResult = mergeWareHourse(maochaoResult);
            replacePrice(maochaoResult,priceMap);
            Collections.sort(maochaoResult,productCompare);
            writeFileMaochao(maochaoResult);
        }

        log.info("生成旗舰店报表");
        String qijian = "d:/goon/qijian.xls";
        List<GoonProduct> qijianList = readXlsData(qijian);
        if (qijianList != null) {
            Map<String,Integer> qijiantransInvMap = readTransInv("GOON-B2C-TM");
            List<GoonProduct> qijianMerge = mergeProduct(qijianList,qijiantransInvMap);
            List<GoonProduct> qijianResult = mergeWareHourse(qijianMerge);
            addTransProduct(qijianResult,qijiantransInvMap,"GOON-B2C-TM");
            qijianResult = mergeWareHourse(qijianResult);
            replacePrice(qijianResult,priceMap);
            Collections.sort(qijianResult,productCompare);
            writeFileQiJian(qijianResult);
        }

        log.info("生成分销报表");
        String fenxiao= "d:/goon/fenxiao.xls";
        List<GoonProduct> fenxiaoList = readXlsData(fenxiao);
        if (fenxiaoList != null) {
            Map<String,Integer> fenxiaotransInvMap = readTransInv("GOON-B2B-FX");
            List<GoonProduct> fenxiaoMerge = mergeProduct(fenxiaoList,fenxiaotransInvMap);
            List<GoonProduct> fenxiaoResult = mergeWareHourse(fenxiaoMerge);
            addTransProduct(fenxiaoResult,fenxiaotransInvMap,"GOON-B2B-FX");
            fenxiaoResult = mergeWareHourse(fenxiaoResult);
            replacePrice(fenxiaoResult,priceMap);
            Collections.sort(fenxiaoResult,productCompare);
            writeFileFenxiao(fenxiaoResult);
        }
    }

    public static List<GoonProduct> readXlsData(String fileName) throws IOException {
        File file = new File(fileName);

        if (!file.exists()) {
            throw new RuntimeException("要制作的xls文件不存在");
        }

        InputStream inputStream = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        int rows = sheet.getPhysicalNumberOfRows();
        List<GoonProduct> list = Lists.newArrayList();
        for (int i  = 1; i < rows;i++) {
            HSSFRow row = sheet.getRow(i);
            GoonProduct product = new GoonProduct();
            product.setSkuName(row.getCell(2).getStringCellValue());
            product.setSkuCode(row.getCell(3).getStringCellValue().trim());
            product.setBrank(row.getCell(7).getStringCellValue());
            String inventory = row.getCell(1).getStringCellValue();
            product.setInventory(StringUtils.isBlank(inventory) ? 0 : Integer.parseInt(inventory));
            product.setWarehourse(row.getCell(8).getStringCellValue().trim());
            product.setPrice(row.getCell(6).getStringCellValue());
            list.add(product);
        }
        return list;
    }

    public static void writeFileZhipin(List<GoonProduct> results) throws IOException {
        File file = new File("d:/goon/out/zhipin-template.xls");

        if (!file.exists()) {
            log.info("纸品模板不存在");
            return;
        }
        POIFSFileSystem fileSystem = new POIFSFileSystem(new FileInputStream(file));
        HSSFWorkbook outWookbook = new HSSFWorkbook(fileSystem);
        HSSFSheet resultSheet = outWookbook.getSheetAt(0);
        for (int i = 0;i < results.size();i++) {
            int rowNum = i +2;
            GoonProduct product = results.get(i);
            HSSFRow row = resultSheet.createRow(rowNum);
            row.createCell(0).setCellValue(product.getSkuCode());
            row.createCell(1).setCellValue(product.getSkuName());
            if (StringUtils.isBlank(product.getPrice())) {
                row.createCell(2).setCellValue(0);
            }
            else {
                row.createCell(2).setCellValue(Integer.parseInt(product.getPrice()));
            }
            if (product.getJashanInv() != 0) {
                row.createCell(3).setCellValue(product.getJashanInv());
            }
            if (product.getJashanTransport() != 0) {
                row.createCell(4).setCellValue(product.getJashanTransport());
            }
        }

        String extendPath = "d:/goon/out/zhipin"+ DateUtils.format(new Date(),"yyyy-MM-dd")+".xls";
        File outFile = new File(extendPath);
        if (outFile.exists()) {
            outFile.delete();
        }
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(outFile);
            outWookbook.write(outputStream);
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
        }
    }

    public static void writeFileMaochao(List<GoonProduct> results) throws IOException {
        File file = new File("d:/goon/out/maochao-template.xls");

        if (!file.exists()) {
            log.info("猫超模板不存在");
            return;
        }
        POIFSFileSystem fileSystem = new POIFSFileSystem(new FileInputStream(file));
        HSSFWorkbook outWookbook = new HSSFWorkbook(fileSystem);
        HSSFSheet resultSheet = outWookbook.getSheetAt(0);
        for (int i = 0;i < results.size();i++) {
            int rowNum = i +2;
            GoonProduct product = results.get(i);
            HSSFRow row = resultSheet.createRow(rowNum);
            row.createCell(0).setCellValue(product.getSkuCode());
            row.createCell(1).setCellValue(product.getSkuName());
            if (StringUtils.isBlank(product.getPrice())) {
                row.createCell(2).setCellValue(0);
            }
            else {
                row.createCell(2).setCellValue(Integer.parseInt(product.getPrice()));
            }
            if (product.getChengduInv() == 0) {
                row.createCell(3);
            }
            else {
                row.createCell(3).setCellValue(product.getChengduInv());
            }
            if (product.getChengduTransport() != 0) {
                row.createCell(4).setCellValue(product.getChengduTransport());
            }
            if (product.getGuangzhouInv() != 0) {
                row.createCell(5).setCellValue(product.getGuangzhouInv());
            }
            if (product.getGuangzhouTrans() != 0) {
                row.createCell(6).setCellValue(product.getGuangzhouTrans());
            }
            if (product.getTianjinLianTInv() != 0) {
                row.createCell(7).setCellValue(product.getTianjinLianTInv());
            }
            if (product.getTianjinLianTTrans() != 0) {
                row.createCell(8).setCellValue(product.getTianjinLianTTrans());
            }
            if (product.getTianjinLianT2Inv() != 0) {
                row.createCell(9).setCellValue(product.getTianjinLianT2Inv());
            }
            if (product.getTianjinLianT2Trans() != 0) {
                row.createCell(10).setCellValue(product.getTianjinLianT2Trans());
            }
            if (product.getTianjinInv() != 0) {
                row.createCell(11).setCellValue(product.getTianjinInv());
            }
            if (product.getTianjinTrans() != 0) {
                row.createCell(12).setCellValue(product.getTianjinTrans());
            }
            if (product.getWuhanInv() != 0) {
                row.createCell(13).setCellValue(product.getWuhanInv());
            }
            if (product.getWuhanTrans() != 0) {
                row.createCell(14).setCellValue(product.getWuhanTrans());
            }
            if (product.getShanghaiInv() != 0) {
                row.createCell(15).setCellValue(product.getShanghaiInv());
            }
            if (product.getShanghaiTrans() != 0) {
                row.createCell(16).setCellValue(product.getShanghaiTrans());
            }
        }
        String extendPath = "d:/goon/out/maochao"+ DateUtils.format(new Date(),"yyyy-MM-dd")+".xls";
        File outFile = new File(extendPath);
        if (outFile.exists()) {
            outFile.delete();
        }
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(outFile);
            outWookbook.write(outputStream);
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
        }
    }

    public static void writeFileQiJian(List<GoonProduct> results) throws IOException {

        File file = new File("d:/goon/out/qijian-template.xls");

        if (!file.exists()) {
            log.info("旗舰模板不存在");
            return;
        }
        POIFSFileSystem fileSystem = new POIFSFileSystem(new FileInputStream(file));
        HSSFWorkbook outWookbook = new HSSFWorkbook(fileSystem);
        HSSFSheet resultSheet = outWookbook.getSheetAt(0);
        for (int i = 0;i < results.size();i++) {
            int rowNum = i +2;
            GoonProduct product = results.get(i);
            HSSFRow row = resultSheet.createRow(rowNum);
            row.createCell(0).setCellValue(product.getSkuCode());
            row.createCell(1).setCellValue(product.getSkuName());
            if (StringUtils.isBlank(product.getPrice())) {
                row.createCell(2).setCellValue(0);
            }
            else {
                row.createCell(2).setCellValue(Integer.parseInt(product.getPrice()));
            }

            if (product.getJashanInv() != 0) {
                row.createCell(3).setCellValue(product.getJashanInv());
            }
            if (product.getJashanTransport() != 0) {
                row.createCell(4).setCellValue(product.getJashanTransport());
            }

            if (product.getTianjinLianTInv() != 0) {
                row.createCell(5).setCellValue(product.getTianjinLianTInv());
            }
            if (product.getTianjinLianTTrans() != 0) {
                row.createCell(6).setCellValue(product.getTianjinLianTTrans());
            }
            if (product.getMaitShanghaiInv() != 0) {
                row.createCell(7).setCellValue(product.getMaitShanghaiInv());
            }
            if (product.getMaitShanghaiTrans() != 0) {
                row.createCell(8).setCellValue(product.getMaitShanghaiTrans());
            }
            if (product.getMaitBeijingInv() != 0) {
                row.createCell(9).setCellValue(product.getMaitBeijingInv());
            }
            if (product.getMaitBeijingTrans() != 0) {
                row.createCell(10).setCellValue(product.getMaitBeijingTrans());
            }
            if (product.getShanghaiInv() != 0) {
                row.createCell(11).setCellValue(product.getShanghaiInv());
            }
            if (product.getShanghaiTrans() != 0) {
                row.createCell(12).setCellValue(product.getShanghaiTrans());
            }
        }
        String extendPath = "d:/goon/out/qijian"+ DateUtils.format(new Date(),"yyyy-MM-dd")+".xls";
        File outFile = new File(extendPath);
        if (outFile.exists()) {
            outFile.delete();
        }
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(outFile);
            outWookbook.write(outputStream);
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
        }
    }

    public static void writeFileFenxiao(List<GoonProduct> results) throws IOException {
        File file = new File("d:/goon/out/fenxiao-template.xls");
        if (!file.exists()) {
            log.info("分销模板不存在");
            return;
        }
        POIFSFileSystem fileSystem = new POIFSFileSystem(new FileInputStream(file));
        HSSFWorkbook outWookbook = new HSSFWorkbook(fileSystem);
        HSSFSheet resultSheet = outWookbook.getSheetAt (0);
        for (int i = 0;i < results.size();i++) {
            int rowNum = i +2;
            GoonProduct product = results.get(i);
            HSSFRow row = resultSheet.createRow(rowNum);
            row.createCell(0).setCellValue(product.getSkuCode());
            row.createCell(1).setCellValue(product.getSkuName());
            if (StringUtils.isBlank(product.getPrice())) {
                row.createCell(2).setCellValue(0);
            }
            else {
                row.createCell(2).setCellValue(Integer.parseInt(product.getPrice()));
            }
            if (product.getChengduInv() != 0) {
                row.createCell(3).setCellValue(product.getChengduInv());
            }
            if (product.getChengduTransport() != 0) {
                row.createCell(4).setCellValue(product.getChengduTransport());
            }
            if (product.getGuangzhouInv() != 0) {
                row.createCell(5).setCellValue(product.getGuangzhouInv());
            }
            if (product.getGuangzhouTrans() != 0) {
                row.createCell(6).setCellValue(product.getGuangzhouTrans());
            }
            if (product.getTianjinLianTInv() != 0) {
                row.createCell(7).setCellValue(product.getTianjinLianTInv());
            }
            if (product.getTianjinLianTTrans() != 0) {
                row.createCell(8).setCellValue(product.getTianjinLianTTrans());
            }
            if (product.getTianjinLianT2Inv() != 0) {
                row.createCell(9).setCellValue(product.getTianjinLianT2Inv());
            }
            if (product.getTianjinLianT2Trans() != 0) {
                row.createCell(10).setCellValue(product.getTianjinLianT2Trans());
            }
            if (product.getTianjinInv() != 0) {
                row.createCell(11).setCellValue(product.getTianjinInv());
            }
            if (product.getTianjinTrans() != 0) {
                row.createCell(12).setCellValue(product.getTianjinTrans());
            }
            if (product.getWuhanInv() != 0) {
                row.createCell(13).setCellValue(product.getWuhanInv());
            }
            if (product.getWuhanTrans() != 0) {
                row.createCell(14).setCellValue(product.getWuhanTrans());
            }
            if (product.getShanghaiInv() != 0) {
                row.createCell(15).setCellValue(product.getShanghaiInv());
            }
            if (product.getShanghaiTrans() != 0) {
                row.createCell(16).setCellValue(product.getShanghaiTrans());
            }
        }
        String extendPath = "d:/goon/out/fenxiao"+ DateUtils.format(new Date(),"yyyy-MM-dd")+".xls";
        File outFile = new File(extendPath);
        if (outFile.exists()) {
            outFile.delete();
        }
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(outFile);
            outWookbook.write(outputStream);
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
        }
    }

    public static List<GoonProduct> mergeProduct(List<GoonProduct> list,Map<String,Integer> transInvMap) {
        log.info("按仓库和skuCode合并库存");
        Map<String,GoonProduct> groupMap = Maps.newHashMap();
        for (GoonProduct product : list) {
            if (StringUtils.isBlank(product.getSkuCode()) ||
                    StringUtils.isBlank(product.getSkuName()) ||
                    StringUtils.isBlank(product.getWarehourse())) {
                continue;
            }
            String key = product.getSkuCode()+"|"+product.getWarehourse();
            GoonProduct temp = groupMap.get(key);
            Integer transInv = transInvMap.get(key);
            if (temp == null) {
                if (StringUtils.equals(product.getWarehourse(),"KEY_JIASHAN")) {
                    product.setJashanInv(product.getInventory());
                    if (transInv != null) {
                        product.setJashanTransport(transInv);
                    }
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_WUHAN")) {
                    product.setWuhanInv(product.getInventory());
                    if (transInv != null) {
                        product.setWuhanTrans(transInv);
                    }
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_CHENGDU")) {
                    product.setChengduInv(product.getInventory());
                    if (transInv != null) {
                        product.setChengduTransport(transInv);
                    }
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_GUANGZHOU")) {
                    product.setGuangzhouInv(product.getInventory());
                    if (transInv != null) {
                        product.setGuangzhouTrans(transInv);
                    }
                }
                else if (StringUtils.equals(product.getWarehourse(),"SH006")) {
                    product.setShanghaiInv(product.getInventory());
                    if (transInv != null) {
                        product.setShanghaiTrans(transInv);
                    }
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_LIANTONG")) {
                    product.setTianjinLianTInv(product.getInventory());
                    if (transInv != null) {
                        product.setTianjinLianTTrans(transInv);
                    }
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_LIANTONG2")) {
                    product.setTianjinLianT2Inv(product.getInventory());
                    if (transInv != null) {
                        product.setTianjinLianT2Trans(transInv);
                    }
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_TIANJING")) {
                    product.setTianjinInv(product.getInventory());
                    if (transInv != null) {
                        product.setTianjinTrans(transInv);
                    }
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_MAITIAN")) {
                    product.setMaitShanghaiInv(product.getInventory());
                    if (transInv != null) {
                        product.setMaitShanghaiTrans(transInv);
                    }
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_MAITIANBJ")) {
                    product.setMaitBeijingInv(product.getInventory());
                    if (transInv != null) {
                        product.setMaitBeijingTrans(transInv);
                    }
                }
                groupMap.put(key,product);
                transInvMap.remove(key);
            } else {
                if (StringUtils.equals(product.getWarehourse(),"KEY_JIASHAN")) {
                    int sum = temp.getJashanInv() + product.getInventory();
                    temp.setJashanInv(sum);
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_WUHAN")) {
                    int sum = temp.getWuhanInv() + product.getInventory();
                    temp.setWuhanInv(sum);
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_CHENGDU")) {
                    int sum = temp.getChengduInv() + product.getInventory();
                    temp.setChengduInv(sum);
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_GUANGZHOU")) {
                    int sum = temp.getGuangzhouInv() + product.getInventory();
                    temp.setGuangzhouInv(sum);
                }
                else if (StringUtils.equals(product.getWarehourse(),"SH006")) {
                    int sum = temp.getShanghaiInv() + product.getInventory();
                    temp.setShanghaiInv(sum);
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_LIANTONG")) {
                    int sum = temp.getTianjinLianTInv() + product.getInventory();
                    temp.setTianjinLianTInv(sum);
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_LIANTONG2")) {
                    int sum = temp.getTianjinLianT2Inv() + product.getInventory();
                    temp.setTianjinLianT2Inv(sum);
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_TIANJING")) {
                    int sum = temp.getTianjinInv() + product.getInventory();
                    temp.setTianjinInv(sum);
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_MAITIAN")) {
                    int sum = temp.getMaitShanghaiInv() + product.getInventory();
                    temp.setMaitShanghaiInv(sum);
                }
                else if (StringUtils.equals(product.getWarehourse(),"KEY_MAITIANBJ")) {
                    int sum = temp.getMaitBeijingInv() + product.getInventory();
                    temp.setMaitBeijingInv(sum);
                }
                else {
                    log.info("未找到匹配项目");
                }
            }
        }

        return new ArrayList<>(groupMap.values());
    }

    public static List<GoonProduct> mergeWareHourse(List<GoonProduct> list) {
        log.info("按skucode合并");
        Map<String,GoonProduct> groupMap = Maps.newHashMap();

        for (GoonProduct product : list) {
            GoonProduct temp = groupMap.get(product.getSkuCode());

            if (temp == null) {
                groupMap.put(product.getSkuCode(),product);
            }
            else {
                if (product.getChengduInv() != 0) {
                    temp.setChengduInv(product.getChengduInv());
                }
                if (product.getChengduTransport() != 0) {
                    temp.setChengduTransport(product.getChengduTransport());
                }
                if (product.getGuangzhouInv() != 0) {
                    temp.setGuangzhouInv(product.getGuangzhouInv());
                }
                if (product.getGuangzhouTrans() != 0) {
                    temp.setGuangzhouTrans(product.getGuangzhouTrans());
                }
                if (product.getTianjinLianTInv() != 0) {
                    temp.setTianjinLianTInv(product.getTianjinLianTInv());
                }
                if (product.getTianjinLianTTrans() != 0) {
                    temp.setTianjinLianTTrans(product.getTianjinLianTTrans());
                }
                if (product.getTianjinLianT2Inv() != 0) {
                    temp.setTianjinLianT2Inv(product.getTianjinLianT2Inv());
                }
                if (product.getTianjinLianT2Trans() != 0) {
                    temp.setTianjinLianT2Trans(product.getTianjinLianT2Trans());
                }
                if (product.getTianjinInv() != 0) {
                    temp.setTianjinInv(product.getTianjinInv());
                }
                if (product.getTianjinTrans() != 0) {
                    temp.setTianjinTrans(product.getTianjinTrans());
                }
                if (product.getWuhanInv() != 0) {
                    temp.setWuhanInv(product.getWuhanInv());
                }
                if (product.getWuhanTrans() != 0) {
                    temp.setWuhanTrans(product.getWuhanTrans());
                }
                if (product.getShanghaiInv() != 0) {
                    temp.setShanghaiInv(product.getShanghaiInv());
                }
                if (product.getShanghaiTrans() != 0) {
                    temp.setShanghaiTrans(product.getShanghaiTrans());
                }
                if (product.getMaitBeijingInv() != 0) {
                    temp.setMaitBeijingInv(product.getMaitBeijingInv());
                }
                if (product.getMaitBeijingTrans() != 0) {
                    temp.setMaitBeijingTrans(product.getMaitBeijingTrans());
                }
                if (product.getMaitShanghaiInv() != 0) {
                    temp.setMaitShanghaiInv(product.getMaitShanghaiInv());
                }
                if (product.getMaitShanghaiTrans() != 0) {
                    temp.setMaitShanghaiTrans(product.getMaitShanghaiTrans());
                }
                if (product.getJashanInv() != 0) {
                    temp.setJashanInv(product.getJashanInv());
                }
                if (product.getJashanTransport() != 0) {
                    temp.setJashanTransport(product.getJashanTransport());
                }
            }
        }
        return new ArrayList<>(groupMap.values());
    }

    public static Map<String,String> readPriceXls() throws IOException {
        String filePath = "d:/goon/priceDetail.xls";
        File file = new File(filePath);
        if (!file.exists()) {
            log.info("价格明细数据不存在");
            throw new RuntimeException("价格明显数据不存在");
        }

        InputStream inputStream = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        Map<String,String> priceMap = Maps.newHashMap();
        for (int i = 0; i < 4 ; i++) {
            HSSFSheet sheet = workbook.getSheetAt(i);
            int rows = sheet.getPhysicalNumberOfRows();
            DecimalFormat decimalFormat = new DecimalFormat("#");
            for (int j=1; j < rows; j++) {
                HSSFRow row = sheet.getRow(j);
                Double skuCode = row.getCell(2).getNumericCellValue();
                Double skuPrice = row.getCell(7).getNumericCellValue();
                priceMap.put(decimalFormat.format(skuCode),decimalFormat.format(skuPrice));
            }
        }
        return priceMap;
    }

    public static void replacePrice(List<GoonProduct> list,Map<String,String> priceMap) {
        for (GoonProduct goonProduct : list) {
            String price =  priceMap.get(goonProduct.getSkuCode());
            goonProduct.setPrice(price);
        }
    }

    public static Map<String,Integer> readTransInv(String type) throws IOException {
        String filePath = "d:/goon/transInv.xls";
        File file = new File(filePath);
        if (!file.exists()) {
            log.info("在途明细数据不存在");
            throw new RuntimeException("在途明细数据不存在");
        }
        InputStream inputStream = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        Map<String,Integer> transInvMap = Maps.newHashMap();
        HSSFSheet sheet = workbook.getSheetAt(0);
        int rows = sheet.getPhysicalNumberOfRows();
        for (int i=1; i<rows; i++) {
            HSSFRow row =  sheet.getRow(i);

            Cell cell = row.getCell(4);
            if (cell == null ||
                    cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                continue;
            }
            String typeCode = cell.getStringCellValue();

            if (!StringUtils.equals(type,typeCode)) {
                continue;
            }

            Cell skuCodeCell = row.getCell(5);
            if (skuCodeCell == null ||
                    skuCodeCell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                skuCodeCell = row.getCell(6);
            }
            String skuCode = skuCodeCell.getStringCellValue();
            String skuName = row.getCell(7).getStringCellValue();

            if (!skuMap.containsKey(skuCode)) {
                skuMap.put(skuCode,skuName);
            }

            Cell wareCell = row.getCell(0);
            if (wareCell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                continue;
            }
            String wareHourse = wareCell.getStringCellValue();

            String dueQty = null;
            int dueCellType = row.getCell(8).getCellType();
            if (Cell.CELL_TYPE_NUMERIC == dueCellType) {
                double qty = row.getCell(8).getNumericCellValue();
                DecimalFormat decimalFormat = new DecimalFormat("##");
                dueQty = decimalFormat.format(qty);
            } else {
                dueQty = row.getCell(8).getStringCellValue();
            }

            String realQty = null;
            if (row.getCell(19) != null) {
                int realQtyType = row.getCell(19).getCellType();
                if (Cell.CELL_TYPE_NUMERIC == realQtyType) {
                    double qty = row.getCell(19).getNumericCellValue();
                    DecimalFormat decimalFormat = new DecimalFormat("##");
                    realQty = decimalFormat.format(qty);
                } else {
                    realQty = row.getCell(19).getStringCellValue();
                }
            }
            Integer dueQtyInt = Integer.parseInt(dueQty);
            Integer realQtyInt = StringUtils.isBlank(realQty) ? 0 : Integer.parseInt(realQty);
            int sub = dueQtyInt - realQtyInt;
            if ( sub > 0) {
                String key = skuCode + "|" + wareHourse;
                if (transInvMap.containsKey(key)) {
                    Integer temp = transInvMap.get(key);
                    Integer sum = temp + sub;
                    transInvMap.put(key,sum);
                }
                else {
                    transInvMap.put(key,sub);
                }
            }
        }

        return transInvMap;
    }


    public static void addTransProduct(List<GoonProduct> list,Map<String,Integer> transMap,String type) {
        if (transMap.isEmpty()) {
            return;
        }

        List<String> transList = new ArrayList<>(transMap.keySet());
        Map<String,GoonProduct> productMap = Maps.newHashMap();
        for (String transKey : transList) {
            String[] transArr = transKey.split("\\|");
            if (transArr.length < 2) {
                continue;
            }
            String skuCode = transArr[0];
            String wareHourseCode = transArr[1];
            GoonProduct product = null;
            if (!productMap.containsKey(skuCode)) {
                product = new GoonProduct();
                product.setSkuCode(skuCode);
                product.setSkuName(skuMap.get(skuCode));
                product.setWarehourse(wareHourseCode);
                productMap.put(skuCode,product);
            } else {
                product = productMap.get(skuCode);
            }

            if (StringUtils.equals(product.getWarehourse(),"KEY_JIASHAN")) {
                product.setJashanTransport(transMap.get(transKey));
            }
            else if (StringUtils.equals(product.getWarehourse(),"KEY_WUHAN")) {
                product.setWuhanTrans(transMap.get(transKey));
            }
            else if (StringUtils.equals(product.getWarehourse(),"KEY_CHENGDU")) {
                product.setChengduTransport(transMap.get(transKey));
            }
            else if (StringUtils.equals(product.getWarehourse(),"KEY_GUANGZHOU")) {
                product.setGuangzhouTrans(transMap.get(transKey));
            }
            else if (StringUtils.equals(product.getWarehourse(),"SH006")) {
                product.setShanghaiTrans(transMap.get(transKey));
            }
            else if (StringUtils.equals(product.getWarehourse(),"KEY_LIANTONG")) {
                product.setTianjinLianTTrans(transMap.get(transKey));
            }
            else if (StringUtils.equals(product.getWarehourse(),"KEY_LIANTONG2")) {
                product.setTianjinLianT2Trans(transMap.get(transKey));
            }
            else if (StringUtils.equals(product.getWarehourse(),"KEY_TIANJING")) {
                product.setTianjinTrans(transMap.get(transKey));
            }
            else if (StringUtils.equals(product.getWarehourse(),"KEY_MAITIAN")) {
                product.setMaitShanghaiTrans(transMap.get(transKey));
            }
            else if (StringUtils.equals(product.getWarehourse(),"KEY_MAITIANBJ")) {
                product.setMaitBeijingTrans(transMap.get(transKey));
            }
        }

        list.addAll(productMap.values());
    }

    public static void copySheet(HSSFWorkbook outWookbook,String filePath) throws Exception {
        Field field = outWookbook.getClass().getDeclaredField("_sheets");
        field.setAccessible(true);
        List<HSSFSheet> sheets = (List<HSSFSheet>) field.get(outWookbook);

        File file = new File(filePath);
        InputStream inputStream = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet copySheet = workbook.getSheetAt(0);
        sheets.add(copySheet);
        outWookbook.setSheetName(sheets.size() - 1, copySheet.getSheetName());
        copySheet.setActive(false);
        outWookbook.setSheetOrder(copySheet.getSheetName(),1);
    }
}
