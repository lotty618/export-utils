package com.quanxi.qxexportutils.util.doc.poi;

import org.python.core.Py;
import org.python.core.PyFunction;
import org.python.core.PyObject;
import org.python.util.PythonInterpreter;

import java.util.Properties;

public class PythonUtils {
    public static void updateToc(String fileName) {
//        System.setProperty("python.home", "lib/jython-2.7.2.jar");

        Properties props = new Properties();
        props.put("python.home", "lib/jython-2.7.2.jar");
        props.put("python.console.encoding", "UTF-8");
        props.put("python.security.respectJavaAccessibility", "false");
        props.put("python.import.site", "false");
        Properties preprops = System.getProperties();
        PythonInterpreter.initialize(preprops, props, new String[0]);

        String pythonFile = "updateTOC.py";
        PythonInterpreter pi = new PythonInterpreter();
        pi.exec("import sys");
        pi.exec("sys.path.append('C:\\\\Program Files\\\\Python\\\\Python39\\\\lib\\\\site-packages')");
        // 加载python程序
        pi.execfile(pythonFile);
        // 调用python中的函数
        PyFunction pf = pi.get("update_toc", PyFunction.class);
        PyObject po = pf.__call__(Py.newString(fileName));
        pi.cleanup();
        pi.close();
    }
}
