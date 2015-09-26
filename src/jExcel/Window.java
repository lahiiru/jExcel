package jExcel;

import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.filechooser.FileFilter;
import javax.swing.table.TableCellEditor;
import javax.xml.stream.XMLStreamException;

/**
 *
 * @author adm.trine@gmail.com
 */
public class Window extends JFrame implements Runnable {

    private Book book;
    private final JToolBar toolbar;
    private final JToolBar toolbar1;
    private final JPanel fbar;
    private final JTextField formulaBar;
    private JLabel barCaption;
    JMenuBar menuBar;
    RibbonBar Ribbon;

    public Window(Book b) {
        this.book = b;
        if (this.book == null) {
            this.book = new Book(this);
        }
        toolbar = new JToolBar();
        toolbar1 = new JToolBar();
        fbar = new JPanel();
        formulaBar = new JTextField(70);
        menuBar = new JMenuBar();
        Ribbon = new RibbonBar(this);

        setName("jExcel");
        setTitle("jExcel");

        addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                closeWindow();
            }
        });

        toolbar.add(fbar);
        formulaBar.setText("Click on a cell");
        formulaBar.setSize(500, 25);
        formulaBar.setFont(new Font("Times", Font.PLAIN, 16));
        formulaBar.setBorder(BorderFactory.createCompoundBorder(formulaBar.getBorder(), BorderFactory.createEmptyBorder(0, 5, 0, 5)));
        formulaBar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                getBook().getSheets().get(getBook().getSheetCollection().getSelectedIndex()).setFromFormulaBar();
            }
        });
        formulaBar.getDocument().addDocumentListener(new DocumentListener() {
            public void changedUpdate(DocumentEvent e) {
                update();
            }

            public void removeUpdate(DocumentEvent e) {
                update();
            }

            public void insertUpdate(DocumentEvent e) {
                update();
            }

            public void update() {
                getBook().getActiveSheet().repaint();
            }
        });
        
        fbar.setComponentOrientation(ComponentOrientation.LEFT_TO_RIGHT);
        fbar.setLayout(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();
        
        //natural height, maximum width
        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridx = 0;
        c.gridy = 0;
        c.ipadx = 0;
        fbar.add(Ribbon,c);
                c.fill = GridBagConstraints.HORIZONTAL;
        c.gridx = 1;
        c.gridy = 0;
        c.ipadx = 0;
        barCaption = new JLabel(" Formula :", SwingConstants.CENTER);
        barCaption.setBorder(BorderFactory.createCompoundBorder(barCaption.getBorder(), BorderFactory.createEmptyBorder(0, 17, 0, 17)));
        fbar.add(barCaption, c);
        c.fill = GridBagConstraints.HORIZONTAL;
        c.weightx = 0.5;
        c.gridx = 2;
        c.gridy = 0;
        fbar.add(formulaBar, c);
        c.fill = GridBagConstraints.HORIZONTAL;
        c.weightx = 0;
        c.gridx = 0;
        c.gridy = 0;
        
        //fbar.add(btn,c);

        JMenu fileMenu = new JMenu("File");
        fileMenu.setMnemonic(KeyEvent.VK_F);
        menuBar.add(fileMenu); // Add the file menu
        JMenuItem newMenuItem = new JMenuItem("New", KeyEvent.VK_N);
        newMenuItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_N, InputEvent.CTRL_MASK));
        newMenuItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                getBook().addNewSheet();
            }
        });
        fileMenu.add(newMenuItem);
        JMenuItem openMenuItem = new JMenuItem("Open", KeyEvent.VK_O);
        openMenuItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_O, InputEvent.CTRL_MASK));
        openMenuItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                openBook();
            }
        });
        fileMenu.add(openMenuItem);
        JMenuItem saveMenuItem = new JMenuItem("Save", KeyEvent.VK_S);
        saveMenuItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_S, InputEvent.CTRL_MASK));
        saveMenuItem.addActionListener(new ActionListener() {

            @Override
            public void actionPerformed(ActionEvent e) {
                saveBook(false);
            }

        });
        fileMenu.add(saveMenuItem);
        JMenuItem saveAsMenuItem = new JMenuItem("Save As", KeyEvent.VK_A);
        saveAsMenuItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_S, InputEvent.CTRL_MASK | InputEvent.SHIFT_MASK));
        saveAsMenuItem.addActionListener(new ActionListener() {

            @Override
            public void actionPerformed(ActionEvent e) {
                saveBook(true);
            }

        });
        fileMenu.add(saveAsMenuItem);
        JMenuItem closeSheetMenuItem = new JMenuItem("Close sheet", KeyEvent.VK_C);
        closeSheetMenuItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                getBook().closeSheet(-1);
            }
        });
        fileMenu.add(closeSheetMenuItem);

        fileMenu.addSeparator();

        JMenuItem closeBookMenuItem = new JMenuItem("Close book", KeyEvent.VK_C);
        closeBookMenuItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                closeBook(0);
            }
        });
        fileMenu.add(closeBookMenuItem);

        JMenuItem exitMenuItem = new JMenuItem("Exit", KeyEvent.VK_E);
        exitMenuItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                closeBook(1);
            }
        });
        fileMenu.add(exitMenuItem);

        JMenu editMenu = new JMenu("Edit");
        editMenu.setMnemonic(KeyEvent.VK_E);

        JMenuItem brushMenuItem = new JMenuItem("Brush", KeyEvent.VK_B);
        brushMenuItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Sheet s = getBook().getSheets().get(getBook().getSheetCollection().getSelectedIndex());
                int c = s.getSelectedColumn();
                int r = s.getSelectedRow();
                s.setBrush(s.getSheetModel().cells[r][c].getRformula());

            }
        });
        editMenu.add(brushMenuItem);
        menuBar.add(editMenu);
        JButton about = new JButton("About");
        about.setBorderPainted(false);
        about.setFocusPainted(false);
        about.setContentAreaFilled(false);
        about.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                showAbout();
            }
        });
        menuBar.add(Box.createHorizontalGlue());
        menuBar.add(about);

        if (this.book.getSheetCollection().getTabRunCount() < 1) {
            formulaBar.setVisible(false);
            barCaption.setVisible(false);
            menuBar.getMenu(0).getItem(2).setEnabled(false);
            menuBar.getMenu(0).getItem(3).setEnabled(false);
            menuBar.getMenu(0).getItem(4).setEnabled(false);
            menuBar.getMenu(0).getItem(6).setEnabled(false);
            ((JButton) Ribbon.getComponents()[2]).setEnabled(false);
            ((JButton) Ribbon.getComponents()[3]).setEnabled(false);
        }

        toolbar.setFloatable(false);
        setJMenuBar(menuBar);
        add(toolbar, BorderLayout.NORTH);

    }

    private void showAbout() {
        JOptionPane.showMessageDialog(this, "jExcel 1.0 \n Developed by: JALP Jayakody.\nadm.trine@gmail.com\n\nAll rights reserved.");
    }

    public void openBook() {
        JFileChooser fc = new JFileChooser();
        fc.setAcceptAllFileFilterUsed(false);
        fc.addChoosableFileFilter(new FileFilter() {
            @Override
            public boolean accept(File f) {
                return f.getPath().endsWith(".xls") | f.getPath().endsWith(".XLS") | f.isDirectory();
            }

            @Override
            public String getDescription() {
                return "jExcel file (.xls)";
            }
        });
        int returnVal = fc.showDialog(this, "Open");
        if (returnVal == 0) {
            String path = fc.getSelectedFile().getPath();
            try {
                if (getBook().getSheetCollection().getTabCount() > 0) {
                    Window w=new Window(null);  
                    Book newBook = new Book(w);
                    XMLPharser xml = new XMLPharser(newBook, path);
                    newBook = xml.getBook();
                    newBook.setSavedPath(path);
                    w.setBook(newBook);
                    (new Thread(w)).start();

                } else {
                    XMLPharser xml = new XMLPharser(getBook(), path);

                    setBook(xml.getBook());
                    getBook().setSavedPath(path);
                    // refresh jtabbed pane
                    getBook().getSheetCollection().revalidate();
                    getBook().getSheetCollection().repaint();

                    getFormulaBar().setEditable(true);
                    getFormulaBar().setText("Click on a cell");

                }
            } catch (XMLStreamException | FileNotFoundException ex) {
                Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

    public void closeWindow() {
        closeBook(1);
    }

    public void closeBook(int exit) {
        boolean confirm = !book.isSaved();
        boolean s = false;
        if (confirm) {
            int n = JOptionPane.showConfirmDialog(
                    this,
                    "Are you sure to save changes?",
                    "jExcel",
                    JOptionPane.YES_NO_CANCEL_OPTION);
            System.out.println("" + n);

            switch (n) {
                case 1:
                    s = true;
                    break;
                case 2:
                    s = false;
                    break;
                case 0:
                    s = saveBook(false);
            }
        } else {
            s = true;
        }
        if (s) {
            try {
                Thread.sleep(20);
            } catch (InterruptedException ex) {
                Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
            }
            if (exit == 1) {
                this.dispose();
            } else {
                int i = 0;
                getBook().getSheetCollection().removeAll();
                ArrayList<Sheet> copy = new ArrayList(getBook().getSheets());
                for (Sheet sheet : copy) {
                    getBook().getSheets().remove(sheet);
                    getBook().setSheetCount(getBook().getSheetCount() - 1);
                }

                formulaBar.setVisible(false);
                barCaption.setVisible(false);
                menuBar.getMenu(0).getItem(2).setEnabled(false);
                menuBar.getMenu(0).getItem(3).setEnabled(false);
                menuBar.getMenu(0).getItem(4).setEnabled(false);
                menuBar.getMenu(0).getItem(6).setEnabled(false);
                ((JButton) Ribbon.getComponents()[2]).setEnabled(false);
                ((JButton) Ribbon.getComponents()[3]).setEnabled(false);
            }
        }
    }

    public boolean saveBook(boolean saveAs) {
        String path;
        if (getBook().getSavedPath().length() < 1 || saveAs) {
            final JFileChooser fc = new JFileChooser();
            int saved;
            fc.setAcceptAllFileFilterUsed(false);

            fc.addChoosableFileFilter(new FileFilter() {
                @Override
                public boolean accept(File f) {
                    return f.getPath().endsWith(".xls") | f.getPath().endsWith(".XLS") | f.isDirectory();
                }

                @Override
                public String getDescription() {
                    return "jExcel file (.xls)";
                }
            });
            saved = fc.showSaveDialog(this);
            if (saved != 0) {
                return false;
            }
            path = fc.getSelectedFile().toString();
            getBook().setSavedPath(path);
        } else {
            path = getBook().getSavedPath();
        }

        if (!path.endsWith(".xls")) {
            path += ".xls";
        }
        System.out.println("" + path);
        try {
            XMLPharser.createXmlFile(this.getBook(), path);
            getBook().setSaved(true);
            return true;
        } catch (IOException ex) {
            Logger.getLogger(Window.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
    }

    @Override
    public void run() {
        add(getBook(), BorderLayout.CENTER);
        setPreferredSize(new Dimension(850, 600));
        pack();
        setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
        setVisible(true);
    }

    /**
     * @return the book
     */
    public Book getBook() {
        return book;
    }

    /**
     * @param book the book to set
     */
    public void setBook(Book book) {
        this.book = book;
    }

    /**
     * @return the toolbar
     */
    public JToolBar getToolbar() {
        return toolbar;
    }

    /**
     * @return the fbar
     */
    public JPanel getFbar() {
        return fbar;
    }

    /**
     * @return the formulaBar
     */
    public JTextField getFormulaBar() {
        return formulaBar;
    }

    /**
     * @return the barCaption
     */
    public JLabel getBarCaption() {
        return barCaption;
    }

    /**
     * @param barCaption the barCaption to set
     */
    public void setBarCaption(JLabel barCaption) {
        this.barCaption = barCaption;
    }

}
