package com.telek.gt;

import static org.junit.Assert.assertTrue;

import com.telek.gt.util.ExeclUtil;
import org.junit.Test;

import java.util.List;
import java.util.Map;

/**
 * Unit test for simple App.
 */
public class AppTest 
{
    /**
     * Rigorous Test :-)
     */
    @Test
   public void testReadExecl(){
        String path = "C:\\Users\\telek\\Desktop\\文件.xlsx";
        List<List<String>> lists = ExeclUtil.readExecl(path);
        System.out.println(lists);
    }

    @Test
    public void testReadExeclByHead(){
        String path = "C:\\Users\\telek\\Desktop\\文件.xlsx";
        List<Map<String, Object>> maps = ExeclUtil.readExeclByHead(path);
        System.out.println(maps);

    }
}
