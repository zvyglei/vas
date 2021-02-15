package org.example.vas.swing;

import org.example.vas.excel.GenerateExcel;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ComponentAdapter;
import java.awt.event.ComponentEvent;
import java.io.File;

/**
 * @author zhao
 * @time 2020/12/6 19:07
 */
public class SwingWindow {
    private JButton btn;
    private JButton btn2;
    private JLabel label;
    private JLabel pathLabel;
    private JLabel msgLabel;

    public void init() {
        // 创建 JFrame 实例
        JFrame frame = new JFrame("VAS");
        try {
            String lookAndFeel =UIManager.getSystemLookAndFeelClassName();
            UIManager.setLookAndFeel(lookAndFeel);
        } catch (Exception e) {
            e.printStackTrace();
        }
        // Setting the width and height of frame
        frame.setSize(350, 200);
        frame.setMinimumSize(new Dimension(350, 200));
        frame.setLocationRelativeTo(null);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        /* 创建面板，这个类似于 HTML 的 div 标签
         * 我们可以创建多个面板并在 JFrame 中指定位置
         * 面板中我们可以添加文本字段，按钮及其他组件。
         */
        JPanel panel = new JPanel();
        // 添加面板
        frame.add(panel);
        /*
         * 调用用户定义的方法并添加组件到面板
         */
        placeComponents(panel);

        // 设置界面可见
        frame.setVisible(true);

        frame.addComponentListener(new ComponentAdapter() {
            @Override
            public void componentResized(ComponentEvent e) {
                pathLabel.setBounds(100, 70, frame.getWidth() - 150, 30);
            }
        });
    }


    private void placeComponents(JPanel panel) {
        /* 布局部分我们这边不多做介绍
         * 这边设置布局为 null
         */
        panel.setLayout(null);
        // 创建按钮
        btn = createBtn("选择文件夹", 18, 30, 120, 33);
        btn.addActionListener(e -> {
            selectFolder(e);
        });
        btn2 = createBtn("生成汇总文件", 150, 30, 120, 33);
        btn2.addActionListener(e -> {
            generateFile(e);
        });

        // 创建 JLabel
        label = createLabel("文件夹路径：", 20, 70, 80, 30);
        pathLabel = createLabel("无", 100, 70, 200, 30);
        pathLabel.setForeground(Color.GRAY);
        msgLabel = createLabel("", 100, 100, 200, 30);

        panel.add(btn);
        panel.add(btn2);
        panel.add(label);
        panel.add(pathLabel);
        panel.add(msgLabel);
    }

    private JButton createBtn(String text, int x, int y, int width, int height) {
        JButton btn = new JButton(text);
        btn.setContentAreaFilled(false);
        // 按钮文本样式
        btn.setFont(new Font("仿宋", Font.TRUETYPE_FONT, 14));
        btn.setBounds(x, y, width, height);
        // 按钮内容与边框距离
        btn.setMargin(new Insets(0, 0, 0, 0));
        return btn;
    }

    private JLabel createLabel(String text, int x, int y, int width, int height) {
        JLabel label = new JLabel(text);
        label.setBounds(x, y, width, height);
        return label;
    }

    private void selectFolder(ActionEvent e) {
        msgLabel.setText("");

        JFileChooser jfc = new JFileChooser();
        jfc.setMultiSelectionEnabled(false);
        jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        jfc.setDialogTitle("请选择文件夹");
        int result = jfc.showOpenDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            File file = jfc.getSelectedFile();
            // String filePath = "<html>" + file.getPath() + "</html>";
            String filePath = file.getPath();
            pathLabel.setText(filePath);
        }
    }

    private void generateFile(ActionEvent e) {
        if ("无".equals(pathLabel.getText())) {
            msgLabel.setForeground(Color.RED);
            msgLabel.setText("未选择文件夹！");
        } else {
            new GenerateExcel().generate(pathLabel.getText());
            msgLabel.setForeground(Color.BLUE);
            msgLabel.setText("已生成汇总文件！");
            try {
                Desktop.getDesktop().open(new File(pathLabel.getText()));
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }
}
