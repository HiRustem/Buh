import java.io.IOException;
import java.util.Scanner;

public class Main {
    private static boolean CYCLE = true;
    private static final String START_MESSAGE = "Введите команду:";
    private static final String UNKNOWN_COMMAND = "Неизвестная команда (для вызова подсказки введите help)";
    private static final String HELP_COMMAND = "help";
    private static final String STOP_COMMAND = "stop";
    private static final String GENERATE_COMMAND = "generate";
    private static final String ADD_SHEET_COMMAND = "addStandardSheet";

    public static void main(String[] args) throws IOException {

        while (CYCLE) {
            Scanner in = new Scanner(System.in);
            System.out.println(START_MESSAGE);
            String string = in.nextLine();

            if (string.equals(HELP_COMMAND)) {
                help();
            } else if (string.equals(STOP_COMMAND)) {
                stop();
            } else if (string.equals(GENERATE_COMMAND)) {
                WorkBook.generateNewService();
            } else if (string.equals(ADD_SHEET_COMMAND)) {
                WorkBook.addServiceSheet();
            } else {
                System.out.println(UNKNOWN_COMMAND);
            }
        }
    }

    public static void help() {
        System.out.println("""
                Команды:\s
                    generate - сформировать служебку\s
                    addStandardSheet - добавить новую страницу в документ\s
                    help - получить полный список команд\s
                    stop - остановить работу приложения""");
    }

    public static void stop() {
        CYCLE = false;
    }
}
