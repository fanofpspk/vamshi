import adding_sum
print(dir(adding_sum))


def test_addition():
    assert adding_sum.add_numbers(3, 4) == 7  # Call the function correctly
    assert adding_sum.add_numbers(-1, 1) == 0
    assert adding_sum.add_numbers(0, 0) == 0

if __name__ == "__main__":
    test_addition()
    print("All tests passed!")

