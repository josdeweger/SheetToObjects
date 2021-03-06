using System;
using FluentAssertions;
using SheetToObjects.Lib.Parsing;
using SheetToObjects.Specs.TestModels;
using Xunit;

namespace SheetToObjects.Specs.Lib.Parsing
{
    public class ParsingStrategyProviderSpecs
    {
        private readonly ParsingStrategyProvider _cellValueParser;

        public ParsingStrategyProviderSpecs()
        {
            _cellValueParser = new ParsingStrategyProvider();
        }

        [Fact]
        public void GivenParsingInt_WhenValueIsNotSet_ResultIsNotValid()
        {
            var result = _cellValueParser.Parse(typeof(int), null);

            result.IsSuccess.Should().BeFalse();
        }

        [Fact]
        public void GivenParsingEnum_WhenValueCanNotBeParsed_ResultIsNotValid()
        {
            var result = _cellValueParser.Parse(typeof(EnumModel), "SomeString");

            result.IsSuccess.Should().BeFalse();
        }

        [Fact]
        public void GivenParsingDouble_WhenValueCanBeParsed_ResultIsValid()
        {
            var doubleValue = 3.3D;

            var result = _cellValueParser.Parse(typeof(double), doubleValue.ToString());

            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be(doubleValue);
        }

        [Fact]
        public void GivenParsingString_WhenValueCanBeParsed_ResultIsValid()
        {
            var stringValue = "MyString";

            var result = _cellValueParser.Parse(typeof(string), stringValue);

            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be(stringValue);
        }

        [Fact]
        public void GivenParsingString_WhenValueCanBeParsedToEnum_ResultIsValid()
        {
            var value = "Second";

            var result = _cellValueParser.Parse(typeof(EnumModel), value);

            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be(EnumModel.Second);
        }

        [Fact]
        public void GivenParsingStringNumber_WhenValueCanBeParsedToEnum_ResultIsValid()
        {
            var value = "2";

            var result = _cellValueParser.Parse(typeof(EnumModel), value);

            result.IsSuccess.Should().BeTrue();
            result.Value.Should().Be(EnumModel.Second);
        }

        [Fact]
        public void GivenParsingStringNumber_WhenValueDoesNotExistInEnum_ResultIsInValid()
        {
            var value = "12";

            var result = _cellValueParser.Parse(typeof(EnumModel), value);

            result.IsSuccess.Should().BeFalse();
        }

        [Fact]
        public void GivenParsingString_WhenValueCantBeParsedToEnum_ResultIsInvalid()
        {
            var value = "notexisting";

            var result = _cellValueParser.Parse(typeof(EnumModel), value);

            result.IsSuccess.Should().BeFalse();
        }

        [Fact]
        public void GivenParsingString_WhenValueIsEmpty_ResultIsInvalid()
        {
            string value = null;

            var result = _cellValueParser.Parse(typeof(EnumModel), value);

            result.IsSuccess.Should().BeFalse();
        }

        [Fact]
        public void GivenParsingString_WhenTypeIsNull_ResultIsInvalid()
        {
            var value = "second";
            var result = _cellValueParser.Parse(null, value);

            result.IsSuccess.Should().BeFalse();
        }

        [Fact]
        public void GivenParsingString_WhenTypeIsNoEnum_ResultIsInvalid()
        {
            var value = "second";
            var result = _cellValueParser.Parse(typeof(double), value);

            result.IsSuccess.Should().BeFalse();
        }

        [Fact]
        public void GivenParsingValidDateTime_WhenFormatIsCorrect_ResultIsValid()
        {
            var format = "yyyy-MM-dd";
            var year = 2018;
            var month = 5;
            var day = 30;
            var value = new DateTime(year, month, day).ToString(format);

            var result = _cellValueParser.Parse(typeof(DateTime), value, format);

            result.IsSuccess.Should().BeTrue();
            ((DateTime) result.Value).Year.Should().Be(year);
            ((DateTime) result.Value).Month.Should().Be(month);
            ((DateTime) result.Value).Day.Should().Be(day);
        }

        [Fact]
        public void GivenParsingDateTime_WhenValueIsEmpty_ResultIsInvalid()
        {
            var value = string.Empty;

            var result = _cellValueParser.Parse(typeof(DateTime), value, "yyyy-MM-dd");

            result.IsSuccess.Should().BeFalse();
        }

        [Fact]
        public void GivenParsingValidDateTime_WhenFormatIsIncorrect_ResultIsInvalid()
        {
            var value = new DateTime(2018, 5, 30).ToString("dd-MM-yyyy");

            var result = _cellValueParser.Parse(typeof(DateTime), value, "yyyy-MM-dd");

            result.IsSuccess.Should().BeFalse();
        }
    }
}
