package com.company.organizer.commons;

public enum NumberConstant {
    ZERO(0), ONE(1),TWO(2), THREE(3), FOUR(4), FIVE(5), SIX(6), SEVEN(7);

    private final Integer number;

    NumberConstant(Integer number) {
        this.number = number;
    }

    public int getNumber() {
        return number;
    }
}
