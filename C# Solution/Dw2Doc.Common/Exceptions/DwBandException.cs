namespace Appeon.DotnetDemo.Dw2Doc.Common.Exceptions
{
    public class DwBandException : Exception
    {
        private string _message;

        public override string Message => _message;


        public DwBandException(string control, string? expectedDwBand, string actualDwBand)
        {
            _message = $"Control {control} expected band {expectedDwBand ?? "<null>"} but got {actualDwBand} instead";
        }
    }
}
