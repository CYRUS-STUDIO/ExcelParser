package com.linchaolong.excelparser;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import com.linchaolong.excelparser.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Excel 导出 lua 文件
 */
public class Excel2Lua {

    public static final String CONFIG_IMPORT = "import";
    public static final String CONFIG_EXPORT = "export";
    public static final String CONFIG_INCLUDE = "include";

    // 根目录，相关文件路径默认基于配置文件所在目录
    private File rootDir;
    // excel 文件路径
    private File importDir;
    // lua 文件导出路径
    private File exportDir;
    // 包含文件内容标记
    private String includeFileFlag;

    public Excel2Lua(String configPath){
        initConfig(configPath);
    }

    /**
     * 创建一个 {@link Excel2Lua} 实例
     *
     * @param configPath    配置文件路径
     * @return
     */
    public static Excel2Lua create(String configPath){
        return new Excel2Lua(configPath);
    }

    /**
     * 遍历excel文件目录，导出lua配置表
     */
    public void run() {
        if (!importDir.exists()) {
            throw new RuntimeException("excel文件目录不存在，请通过" + CONFIG_IMPORT + "字段配置。");
        }

        // 导出所有excel文件
        for (File file : importDir.listFiles()) {
            if (ExcelUtils.isExcel(file.getName())) {
                excel2Lua(file.getPath());
            } else {
                System.err.println(String.format("'%s' is not a excel file!!!", file.getName()));
            }
        }
    }

    /**
     * Excel 导出 lua
     *
     * @param excelPath excel 文件路径
     */
    private void excel2Lua(String excelPath) {

        // 检查文件是否存在
        File excelFile = new File(excelPath);
        if (!excelFile.exists()) {
            throw new RuntimeException(excelPath + " ，文件不存在。");
        }

        // 初始化导出目录
        if (!exportDir.exists()) {
            exportDir.mkdirs();
        }

        // excel表对象
        Workbook workbook = ExcelUtils.workbook(excelPath);

        // 获取第1页的表格，索引从0开始
        Sheet sheet = workbook.getSheetAt(0);

        // 获取总行数
        int totalRow = sheet.getLastRowNum();
        System.out.println(excelFile.getName() + " 表格数量：" + workbook.getNumberOfSheets() + " 表格名称：" + sheet.getSheetName() + " 行数：" + sheet.getLastRowNum());

        // 第一行：字段名称
        Row keyRow = sheet.getRow(0);
        // 第二行：描述
        // Row descRow = sheet.getRow(1);

        String excelFileName = excelFile.getName().substring(0,
                excelFile.getName().lastIndexOf('.'));
        // lua文件
        File luaFile = new File(exportDir, excelFileName + ".lua");
        if (luaFile.exists()) {
            luaFile.delete();
        }

        try {
            BufferedWriter out = new BufferedWriter(new FileWriter(luaFile));
            out.append(excelFileName + " = {");
            out.newLine();

            System.out.println("开始导出：" + luaFile.getName());
            // 迭代每一行
            Row row = null;
            for (int i = 2; i <= totalRow; i++) {
                // 获取每一行
                row = sheet.getRow(i);
                if(row == null){
                    continue;   // 跳过空行
                }
                // 获取当前行的列数
                // cellCount = row.getLastCellNum();
                System.out.print(i + " : ");

                Cell keyCell = row.getCell(0);
                out.append("[" + ExcelUtils.intValue(keyCell) + "] = {");
                out.newLine();

                // 迭代每一列
                for (int j = row.getFirstCellNum() + 1; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    // 忽略空单元格
                    if ("null".equals(cell + "")) {
                        // System.out.println("有空单元格");
                    } else {
                        System.out.print(cell + "\t");
                        out.append(keyRow.getCell(j).toString()).append(" = ")
                                .append(ExcelUtils.stringValue(cell, includeFileFlag, rootDir)).append(",");
                        out.newLine();
                    }
                }
                out.append("},");
                out.newLine();

                System.out.println();
                out.append(",");
                out.newLine();
            }
            out.append("}");
            out.close();
        } catch (IOException e) {
            throw new RuntimeException(luaFile.getPath() + "，文件导出失败。");
        }
        System.out.println("导出完成：" + luaFile.getName());
    }

    /**
     * 初始化配置
     *
     * @param configPath
     */
    private void initConfig(String configPath) {
        File file = new File(configPath);
        if (!file.exists()) {
            throw new RuntimeException("配置文件 '"+file.getAbsolutePath()+"' 不存在，初始化配置失败。");
        }

        rootDir = file.getParentFile();
        Map<String, String> configMap = new HashMap<>();
        try (BufferedReader in = new BufferedReader(new FileReader(file))){
            System.out.println("初始化配置。");
            String line;
            while ((line = in.readLine()) != null) {
                if (line.contains("=")) {
                    String kv[] = line.split("=");
                    configMap.put(kv[0].trim(), kv[1].trim());
                }
            }

            includeFileFlag = configMap.get(CONFIG_INCLUDE);
            importDir = new File(rootDir, configMap.get(CONFIG_IMPORT));
            exportDir = new File(rootDir, configMap.get(CONFIG_EXPORT));

            for (Entry<String, String> entry : configMap.entrySet()) {
                System.out.println(entry.getKey() + " = " + entry.getValue());
            }
            System.out.println("配置初始化完成。");
        } catch (FileNotFoundException e) {
            throw new RuntimeException("配置文件" + configPath + "不存在，初始化配置失败。");
        } catch (IOException e) {
            throw new RuntimeException("读取配置文件" + configPath + "失败。");
        }
    }

    public static void main(String[] args) {
//        Excel2Lua.create("./Excel2Lua/config.txt").run();
        Excel2Lua.create(args[0]).run(); // 从命令行参数读取文件路径
    }

}
