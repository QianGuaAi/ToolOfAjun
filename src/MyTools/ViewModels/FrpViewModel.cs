using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using MyTools.Services;

namespace MyTools.ViewModels
{
    public sealed class FrpViewModel : INotifyPropertyChanged, IDisposable
    {
        private readonly MainViewModel _owner;
        private readonly FrpProcessManager _manager = new FrpProcessManager();
        private string _frpServerAddress = string.Empty;
        private int _frpServerPort = 7000;
        private string _frpToken = string.Empty;
        private string _statusHint = "填写服务器信息和端口规则后启动隧道。";
        private FrpTunnelRule _draftRule;
        private bool _isLoadingConfig;
        private string _clientId = string.Empty;
        private bool _disposed;

        public FrpViewModel(MainViewModel owner)
        {
            _owner = owner;
            FrpRules = new ObservableCollection<FrpTunnelRule>();
            FrpRules.CollectionChanged += FrpRules_OnCollectionChanged;
            DraftRule = new FrpTunnelRule { Type = "tcp", IsEnabled = true };

            StartTunnelCommand = new AsyncRelayCommand(StartTunnelAsync, () => CanStart);
            StopTunnelCommand = new RelayCommand(StopTunnel, () => CanStop);
            AddRuleCommand = new RelayCommand(AddRule, CanAddDraftRule);
            RemoveRuleCommand = new RelayParameterCommand(RemoveRule, parameter => parameter is FrpTunnelRule);
            SaveConfigCommand = new AsyncRelayCommand(SaveConfigAsync, () => !_isLoadingConfig);
            LoadConfigCommand = new AsyncRelayCommand(LoadConfigAsync, () => !_isLoadingConfig);

            _manager.StateChanged += Manager_OnStateChanged;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public string FrpServerAddress
        {
            get => _frpServerAddress;
            set
            {
                var next = value ?? string.Empty;
                if (_frpServerAddress == next)
                {
                    return;
                }

                _frpServerAddress = next;
                OnPropertyChanged();
                NotifyStateProperties();
            }
        }

        public int FrpServerPort
        {
            get => _frpServerPort;
            set
            {
                if (_frpServerPort == value)
                {
                    return;
                }

                _frpServerPort = value;
                OnPropertyChanged();
                NotifyStateProperties();
            }
        }

        public string FrpToken
        {
            get => _frpToken;
            set
            {
                var next = value ?? string.Empty;
                if (_frpToken == next)
                {
                    return;
                }

                _frpToken = next;
                OnPropertyChanged();
                NotifyStateProperties();
            }
        }

        public string StatusHint
        {
            get => _statusHint;
            private set
            {
                var next = value ?? string.Empty;
                if (_statusHint == next)
                {
                    return;
                }

                _statusHint = next;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<FrpTunnelRule> FrpRules { get; }

        public FrpTunnelRule DraftRule
        {
            get => _draftRule;
            set
            {
                if (ReferenceEquals(_draftRule, value))
                {
                    return;
                }

                if (_draftRule != null)
                {
                    _draftRule.PropertyChanged -= DraftRule_OnPropertyChanged;
                }

                _draftRule = value ?? new FrpTunnelRule { Type = "tcp", IsEnabled = true };
                _draftRule.PropertyChanged += DraftRule_OnPropertyChanged;
                OnPropertyChanged();
                NotifyStateProperties();
            }
        }

        public FrpState ConnectionState => _manager.State;
        public string ConnectionStatusText => _manager.StatusMessage;
        public bool IsRunning => _manager.State == FrpState.Starting || _manager.State == FrpState.Running;
        public bool CanStart => !IsRunning && HasRequiredConfig && HasEnabledRules;
        public bool CanStop => IsRunning;
        public bool HasEnabledRules => FrpRules.Any(rule => rule.IsEnabled);
        public bool HasRules => FrpRules.Count > 0;

        public bool HasRequiredConfig =>
            !string.IsNullOrWhiteSpace(FrpServerAddress)
            && FrpServerPort >= 1
            && FrpServerPort <= 65535
            && !string.IsNullOrWhiteSpace(FrpToken);

        public string PublicAddressPreview
        {
            get
            {
                var enabledRules = FrpRules.Where(rule => rule.IsEnabled).ToList();
                if (enabledRules.Count == 0)
                {
                    return "未添加隧道规则";
                }

                if (enabledRules.Count == 1)
                {
                    var address = string.IsNullOrWhiteSpace(FrpServerAddress) ? "服务器地址" : FrpServerAddress.Trim();
                    return $"{address}:{enabledRules[0].RemotePort}";
                }

                return $"已启用 {enabledRules.Count} 条隧道";
            }
        }

        public ICommand StartTunnelCommand { get; }
        public ICommand StopTunnelCommand { get; }
        public ICommand AddRuleCommand { get; }
        public ICommand RemoveRuleCommand { get; }
        public ICommand SaveConfigCommand { get; }
        public ICommand LoadConfigCommand { get; }

        public async Task LoadConfigAsync()
        {
            if (_isLoadingConfig)
            {
                return;
            }

            _isLoadingConfig = true;
            CommandManager.InvalidateRequerySuggested();

            try
            {
                var config = await FrpService.LoadConfigAsync();
                _clientId = string.IsNullOrWhiteSpace(config.ClientId) ? Guid.NewGuid().ToString("N") : config.ClientId;
                FrpServerAddress = config.ServerAddress ?? string.Empty;
                FrpServerPort = FrpService.IsValidPort(config.ServerPort) ? config.ServerPort : 7000;
                FrpToken = FrpService.DecryptToken(config.EncryptedToken);

                foreach (var rule in FrpRules)
                {
                    rule.PropertyChanged -= FrpRule_OnPropertyChanged;
                }

                FrpRules.Clear();

                var rules = await FrpService.LoadRulesAsync();
                foreach (var rule in rules)
                {
                    rule.Type = "tcp";
                    rule.PropertyChanged += FrpRule_OnPropertyChanged;
                    FrpRules.Add(rule);
                }

                StatusHint = HasRules ? "配置已加载。" : "填写服务器信息和端口规则后启动隧道。";
            }
            finally
            {
                _isLoadingConfig = false;
                NotifyAll();
            }
        }

        public async Task SaveConfigAsync()
        {
            ValidateConfigForSave();

            if (string.IsNullOrWhiteSpace(_clientId))
            {
                _clientId = Guid.NewGuid().ToString("N");
            }

            var config = new FrpServerConfig
            {
                ServerAddress = FrpServerAddress.Trim(),
                ServerPort = FrpServerPort,
                EncryptedToken = FrpService.EncryptToken(FrpToken),
                ClientId = _clientId
            };

            await FrpService.SaveConfigAsync(config);
            await FrpService.SaveRulesAsync(FrpRules);

            StatusHint = "配置已保存。";
            AppLogService.Information(
                "FRP config saved: {Server}:{Port}, rules={RuleCount}",
                config.ServerAddress,
                config.ServerPort,
                FrpRules.Count);
            NotifyAll();
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _manager.StateChanged -= Manager_OnStateChanged;
            FrpRules.CollectionChanged -= FrpRules_OnCollectionChanged;
            foreach (var rule in FrpRules)
            {
                rule.PropertyChanged -= FrpRule_OnPropertyChanged;
            }

            if (_draftRule != null)
            {
                _draftRule.PropertyChanged -= DraftRule_OnPropertyChanged;
            }

            _manager.Dispose();
        }

        private async Task StartTunnelAsync()
        {
            await SaveConfigAsync();

            var frpcExePath = await FrpService.EnsureFrpcExtractedAsync();
            var config = new FrpServerConfig
            {
                ServerAddress = FrpServerAddress.Trim(),
                ServerPort = FrpServerPort,
                EncryptedToken = string.Empty,
                ClientId = _clientId
            };

            var ini = FrpService.BuildFrpcIni(config, FrpToken, FrpRules);
            await _manager.StartAsync(frpcExePath, ini);

            StatusHint = _manager.State == FrpState.Error ? _manager.StatusMessage : "隧道启动命令已发送。";
            NotifyAll();
        }

        private void StopTunnel()
        {
            _manager.Stop();
            StatusHint = "隧道已停止。";
            NotifyAll();
        }

        private void AddRule()
        {
            if (!CanAddDraftRule())
            {
                StatusHint = "请检查本机端口、远程端口，且远程端口不能重复。";
                NotifyAll();
                return;
            }

            var rule = new FrpTunnelRule
            {
                Name = DraftRule.Name,
                Type = "tcp",
                LocalPort = DraftRule.LocalPort,
                RemotePort = DraftRule.RemotePort,
                Description = DraftRule.Description,
                IsEnabled = true
            };

            rule.PropertyChanged += FrpRule_OnPropertyChanged;
            FrpRules.Add(rule);
            DraftRule = new FrpTunnelRule { Type = "tcp", IsEnabled = true };
            StatusHint = "规则已添加。请确认服务器安全组已放行远程端口。";
            NotifyAll();
        }

        private void RemoveRule(object parameter)
        {
            if (!(parameter is FrpTunnelRule rule))
            {
                return;
            }

            rule.PropertyChanged -= FrpRule_OnPropertyChanged;
            FrpRules.Remove(rule);
            StatusHint = "规则已删除。";
            NotifyAll();
        }

        private bool CanAddDraftRule()
        {
            return DraftRule != null
                && FrpService.IsValidPort(DraftRule.LocalPort)
                && FrpService.IsValidPort(DraftRule.RemotePort)
                && FrpRules.All(rule => rule.RemotePort != DraftRule.RemotePort);
        }

        private void ValidateConfigForSave()
        {
            if (string.IsNullOrWhiteSpace(FrpServerAddress))
            {
                StatusHint = "请填写 frp 服务器地址。";
                throw new InvalidOperationException(StatusHint);
            }

            if (!FrpService.IsValidPort(FrpServerPort))
            {
                StatusHint = "frp 服务器端口必须在 1-65535。";
                throw new InvalidOperationException(StatusHint);
            }

            if (string.IsNullOrWhiteSpace(FrpToken))
            {
                StatusHint = "请填写 frp Token。";
                throw new InvalidOperationException(StatusHint);
            }
        }

        private void Manager_OnStateChanged(object sender, EventArgs e)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
            {
                dispatcher.BeginInvoke(new Action(RefreshManagerState));
                return;
            }

            RefreshManagerState();
        }

        private void RefreshManagerState()
        {
            if (_manager.State == FrpState.Running)
            {
                StatusHint = "隧道正在运行。请确认本机服务已启动，且服务器安全组已放行远程端口。";
            }
            else if (_manager.State == FrpState.Error)
            {
                StatusHint = _manager.StatusMessage;
            }

            NotifyAll();
        }

        private void FrpRules_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (FrpTunnelRule rule in e.OldItems)
                {
                    rule.PropertyChanged -= FrpRule_OnPropertyChanged;
                }
            }

            if (e.NewItems != null)
            {
                foreach (FrpTunnelRule rule in e.NewItems)
                {
                    rule.PropertyChanged -= FrpRule_OnPropertyChanged;
                    rule.PropertyChanged += FrpRule_OnPropertyChanged;
                }
            }

            NotifyAll();
        }

        private void FrpRule_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            NotifyAll();
        }

        private void DraftRule_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            NotifyStateProperties();
        }

        private void NotifyAll()
        {
            OnPropertyChanged(nameof(ConnectionState));
            OnPropertyChanged(nameof(ConnectionStatusText));
            NotifyStateProperties();
        }

        private void NotifyStateProperties()
        {
            OnPropertyChanged(nameof(IsRunning));
            OnPropertyChanged(nameof(CanStart));
            OnPropertyChanged(nameof(CanStop));
            OnPropertyChanged(nameof(HasEnabledRules));
            OnPropertyChanged(nameof(HasRules));
            OnPropertyChanged(nameof(HasRequiredConfig));
            OnPropertyChanged(nameof(PublicAddressPreview));
            CommandManager.InvalidateRequerySuggested();
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
