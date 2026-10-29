#include <stdio.h>

// A simple C program for your lab record.
int main(void) {
    int marks[] = {85, 92, 78, 90};
    int count = sizeof(marks) / sizeof(marks[0]);
    int total = 0;

    for (int i = 0; i < count; i++) {
        total += marks[i];
    }

    printf("Average: %.2f\n", (double)total / count);
    return 0;
}
