using System.ComponentModel;
using System.Text;

namespace SysmacSimulator
{
    //为了确保 DataGridView 能够实时更新数据，我们需要确保 SimulatorVariable 类实现了 INotifyPropertyChanged 接口，并且在更新 Value 属性时触发属性变化通知。
    //这样 BindingList<T> 才能检测到属性的变化并通知 DataGridView 更新视图。
    public class SimulatorVariable : INotifyPropertyChanged
    {
        private string _variableName;
        private byte[] _revision;
        private byte[] _address;
        private int _size;
        private string _type;
        private int? _lowIndex;
        private int? _highIndex;
        private object _value;

        public string VariableName
        {
            get => _variableName;
            set
            {
                if (_variableName != value)
                {
                    _variableName = value;
                    OnPropertyChanged(nameof(VariableName));
                }
            }
        }

        public string RevisionString
        {
            get => ByteArrayToString(_revision);
        }

        public string AddressString
        {
            get => ByteArrayToString(_address);
        }

        public byte[] Revision
        {
            get => _revision;
            set
            {
                if (_revision != value)
                {
                    _revision = value;
                    OnPropertyChanged(nameof(Revision));
                    OnPropertyChanged(nameof(RevisionString));
                }
            }
        }

        public byte[] Address
        {
            get => _address;
            set
            {
                if (_address != value)
                {
                    _address = value;
                    OnPropertyChanged(nameof(Address));
                    OnPropertyChanged(nameof(AddressString));
                }
            }
        }

        public int Size
        {
            get => _size;
            set
            {
                if (_size != value)
                {
                    _size = value;
                    OnPropertyChanged(nameof(Size));
                }
            }
        }

        public string Type
        {
            get => _type;
            set
            {
                if (_type != value)
                {
                    _type = value;
                    OnPropertyChanged(nameof(Type));
                }
            }
        }

        public int? LowIndex
        {
            get => _lowIndex;
            set
            {
                if (_lowIndex != value)
                {
                    _lowIndex = value;
                    OnPropertyChanged(nameof(LowIndex));
                }
            }
        }

        public int? HighIndex
        {
            get => _highIndex;
            set
            {
                if (_highIndex != value)
                {
                    _highIndex = value;
                    OnPropertyChanged(nameof(HighIndex));
                }
            }
        }

        public object Value
        {
            get => _value;
            set
            {
                if (_value != value)
                {
                    _value = value;
                    OnPropertyChanged(nameof(Value));
                }
            }
        }

        public SimulatorVariable(string variableName, byte[] revision, byte[] address, int size, object value)
        {
            VariableName = variableName;
            Revision = revision;
            Address = address;
            Size = size;
            Type = null;
            LowIndex = null;
            HighIndex = null;
            Value = value;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string ByteArrayToString(byte[] bytes)
        {
            if (bytes == null) return string.Empty;
            return Encoding.UTF8.GetString(bytes);
        }
    }
}