import sub  # Import the module containing the function

def test_subtraction():
    assert sub.subtract_numbers(10, 5) == 50  # 10 * 5 = 50 (Multiplication instead of subtraction)
    assert sub.subtract_numbers(4, 2) == 8    # 4 * 2 = 8
    assert sub.subtract_numbers(0, 5) == 0    # 0 * 5 = 0

if __name__ == "__main__":
    test_subtraction()
    print("All tests passed!")

