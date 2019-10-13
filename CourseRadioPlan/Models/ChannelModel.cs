using System;
namespace CourseRadioPlan.Models
{
    public class ChannelModel
    {
        public string Position { get; set; }
        public string Type { get; set; }
        public string Number { get; set; }

        public override int GetHashCode()
        {
            int result = 0;
            if (this.Position != null)
            {
                result += this.Position.GetHashCode();
            }
            if (this.Type != null)
            {
                result += this.Type.GetHashCode();
            }
            if (this.Number != null)
            {
                result += this.Number.GetHashCode();
            }
            return result;
        }

        public override bool Equals(object obj)
        {
            var that = obj as ChannelModel;
            if (that == null)
            {
                return false;
            }
            return this.Position == that.Position &&
                this.Type == that.Type &&
                this.Number == that.Number;
        }
    }
}
