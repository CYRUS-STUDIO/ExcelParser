package com.linchaolong.excelparser.utils;

import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * 打印工具类
 *
 * Created by linchaolong on 2016/12/2.
 */

public class PrintUtils {

  public static final String TAG = "PrintUtils";

  public static void file(File file){
    try {
      if(!file.exists()){
        return;
      }
      if(file.isDirectory()){
        array(file.list());
      }else{
        print(FileUtils.readFileToString(file));
      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  public static <T> void array(T[] array){
    for(T t : array){
      print(t.toString());
    }
  }

  public static <T> void list(List<T> list){
    for(T t : list){
      print(t.toString());
    }
  }

  public static void print(String str){
    System.out.println(str);
  }

}
