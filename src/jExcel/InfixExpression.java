package jExcel;

import java.util.Stack;

/**
 *
 * @author adm.trine@gmail.com
 */
public class InfixExpression {

    private String expression = "";
    private Stack postfixEx;
    private String solvedEx = "";

    public InfixExpression(String input) {
        expression = "(" + input + ")";
    }

    private int getPrecedence(String operator) {
        String[] precedence = {"+-", "*/"};
        int p = 0;
        for (String i : precedence) {
            p++;
            if (i.contains(operator)) {
                return p;
            }
        }
        return -1;
    }

    private void printStack(Stack s) {
        Stack reversed = new Stack();
        for (Object l : s.toArray()) {
            reversed.insertElementAt(l, 0);
        }
        for (Object k : reversed.toArray()) {
            System.out.print(k.toString());
        }
    }

    private Stack getStack() {
        String input = expression;
        //*********//
        String r[] = {"*-:*~", "--:+", "-+:-", "++:+", "+-:-", "/-:/~", "(-:(~", ")-:)~"};
        for (String i : r) {
            String[] v = i.split(":");
            input = input.replace(v[0], v[1]);
        }
        //**********//
        Stack tokens = new Stack();
        String buffer = "";
        for (int i = 0; i < input.length(); i++) {
            if (input.substring(i, i + 1).matches("[0-9\\.~]") && i < input.length() - 1) {
                buffer += input.substring(i, i + 1);
            } else {
                if (buffer.length() > 0) {

                    tokens.push(buffer);
                }
                //System.out.println("kk"+buffer);
                buffer = "";
                tokens.push(input.substring(i, i + 1));
                //System.out.println(input.substring(i, i+1));
            }
        }
        Stack reversed = new Stack();
        for (Object l : tokens.toArray()) {
            reversed.insertElementAt(l, 0);
        }

        return reversed;
    }

    public String convertToPostfix() {
        Stack tokens = getStack();
        Stack postFix = new Stack();
        Stack operator = new Stack();
        for (Object x : tokens.toArray()) {
            Object i = tokens.peek();
            if (i.toString().matches("(~)?[0-9\\.]+")) {
                postFix.push(tokens.pop());
            } else if (i.toString().matches("[(]")) {
                operator.push(tokens.pop());
            } else if (i.toString().matches("[)]")) {
                tokens.pop();
                for (Object xx : operator.toArray()) {
                    Object n = operator.peek();
                    if (n.toString().matches("[(]")) {
                        operator.pop();
                        break;
                    }
                    postFix.push(operator.pop());
                }

            } else {
                while (!operator.isEmpty() && !(operator.peek().toString().matches("[(]"))) {
                    if (getPrecedence(operator.peek().toString()) >= getPrecedence(i.toString())) {
                        //System.out.println("\n pushed to post"+operator.peek().toString());
                        postFix.push(operator.pop());
                    } else {
                        break;
                    }
                }
                operator.push(tokens.pop());
            }
            if (tokens.empty()) {
                for (Object xx : operator.toArray()) {
                    postFix.push(operator.pop());
                }
            }
            /* debug ....
             System.out.println("\npostFix");
             System.out.println("\npostFix");
             printStack(postFix);
             System.out.println("\noperator");
             printStack(operator);
             System.out.println("\ntokens");
             printStack(tokens);
             System.out.println("\n=================");
             */
        }
        String r = "";
        for (Object l : postFix.toArray()) {
            r += l.toString();
        }
        //printStack(postFix);
        postfixEx = postFix;
        return r;
    }

    public double evaluateEx() {
        convertToPostfix();
        Stack value;
        value = postfixEx;
        for (Object i : value.toArray()) {
            if (i.toString().matches("[\\+\\-\\*\\/]")) {
                int idx = value.indexOf(i);
                double res = 0.0;
                String in2 = value.get(idx - 1).toString().replace("~", "-");
                String in1 = value.get(idx - 2).toString().replace("~", "-");;
                if (i.toString().matches("[*]")) {
                    res = Double.valueOf(in1) * Double.valueOf(in2);
                }
                if (i.toString().matches("[+]")) {
                    res = Double.valueOf(in1) + Double.valueOf(in2);
                }
                if (i.toString().matches("[-]")) {
                    res = Double.valueOf(in1) - Double.valueOf(in2);
                }
                if (i.toString().matches("[/]")) {
                    res = Double.valueOf(in1) / Double.valueOf(in2);
                }
                value.remove(idx - 2);
                value.remove(idx - 2);
                value.remove(idx - 2);
                value.add(idx - 2, String.valueOf(res));
            }
        }
        solvedEx = value.peek().toString();
        return Double.valueOf(solvedEx);
    }

    public String getValue() {
        evaluateEx();
        return solvedEx;
    }
}
