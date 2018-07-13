package tests;

public class Test implements Cloneable {
    private int a;
    public static void main(String[] args) {
        Test test=new Test();
        test.setA(1);
        Test test1=test.clone();
        System.out.println(test1.getA());
    }

    public int getA() {
        return a;
    }

    public void setA(int a) {
        this.a = a;
    }

    public Test clone()

    {

        Object obj = null;

        try

        {

            obj = super.clone();

            return (Test) obj;

        }

        catch(CloneNotSupportedException e)

        {

            System.out.println("不支持复制！");

            return null;

        }

    }
}
