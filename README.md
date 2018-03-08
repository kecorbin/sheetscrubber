# Sheet Scrubber

A python utility to cleanup large spreadsheets.

This utility is spreadsheet specific, but should be able to be easily modified to
your use case.

## Scenario

Ever received a spreadsheet which contains 500 columnns, and 400k rows?
Even opening such a spreadsheet can be a pain in the !@##! and we are likely only
interested in a small subset of that data.

## Solution

Let's create our own version of the spreadsheet with the following criteria

* Eliminate rows that don't have interesting data by checking for a condition on the
value of a specific column in that row (in our case <0)

* Size column width to match the largest value in that column

* Eliminate columns that contain extra data we don't care about

* Add an additional column containing a formula to apply to that rows

* Add some additional data (average of a column in our case)


## Installation

1. Clone this repository

  git clone https://github.com/kecorbin/sheetscrubber
  cd sheetscrubber

2. Run the installation script

  ./install.sh
