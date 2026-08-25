# Student Grade System
# CAIE Course Program - Social Eagle

def calculate_grade(mark):
    """Return the letter grade for a valid mark."""
    if mark >= 90:
        return "A"
    elif mark >= 80:
        return "B"
    elif mark >= 70:
        return "C"
    elif mark >= 60:
        return "D"
    else:
        return "E"


def main():
    try:
        #Asks the user to enter a mark. Converts the input from text to a decimal number using float().
        mark = float(input("Enter your mark (0-100): "))

        # Validate the mark range - Checks that the mark is between 0 and 100.
        if mark < 0 or mark > 100:
            print("Invalid mark. Please enter a number between 0 and 100.")
            return
        #Calls calculate_grade(mark).
        grade = calculate_grade(mark)

        # Display the result - Prints the mark and resulting grade.
        print(f"Mark: {mark:g} -> Grade: {grade}")
    #Handles invalid non-numeric input using except ValueError.
    except ValueError:
        print("Invalid input. Please enter a valid number.")

#ensures main() runs only when this file is executed directly, not when it is imported into another Python file.
if __name__ == "__main__":
    main()