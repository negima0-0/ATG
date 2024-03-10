import csv
import json
from netmiko import ConnectHandler
import configparser

# ユーザ名とパスワードを取得する関数
def get_credentials():
    config = configparser.ConfigParser()
    config.read('config.ini')

    username = config.get('credentials', 'username')
    password = config.get('credentials', 'password')

    return username, password

# ユーザ名とパスワードを取得
username, password = get_credentials()

# CSVファイルからホスト名とIPアドレスの対応を読み込む
hosts = []
with open('hosts.csv', 'r') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if 'hostname' in row and 'ip_address' in row:
            hosts.append((row['hostname'], row['ip_address']))
        elif 'hostname' in row:
            hosts.append((row['hostname'], None))
        elif 'ip_address' in row:
            hosts.append((None, row['ip_address']))

proxy_ssh_info = {
    'device_type': 'juniper_junos',
    'ip': '10.10.0.51',
    'username': 'negima',
    'password': 'raindrop3',
}

# ホストへの接続
for hostname, ip_address in hosts:
    try:
        # ホストへのSSH接続（ホスト名での接続）
        with ConnectHandler(
            device_type='juniper_junos',
            ip=ip_address,
            hostname=hostname,
            username=username,
            password=password,
            port=22,  # SSHポート（デフォルトは22）
            global_delay_factor=2,  # コマンド実行後のディレイ係数
            timeout=30,  # レスポンスまでのタイムアウトを30秒に設定(デフォルトは5秒)
            session_log=f"session_logs/{hostname or ip_address}.log",
            proxy=proxy_ssh_info,
        ) as ssh:
            # コマンドを実行してJSONデータを取得
            json_output = ssh.send_command('show interfaces extensive ge-0/0/0 | display json | no-more')

    except Exception as e:
        # ホスト名での接続が失敗した場合、IPアドレスで再接続を試みる
        print(f"ホスト名での接続に失敗しました。IPアドレス {ip_address} を使用して再接続を試みます...")
        try:
            # ホストへのSSH接続（IPアドレスでの接続）
            with ConnectHandler(
                device_type='juniper_junos',
                ip=ip_address,
                username=username,
                password=password,
                port=22,  # SSHポート（デフォルトは22）
                global_delay_factor=2,  # コマンド実行後のディレイ係数
                timeout = 30,  # レスポンスまでのタイムアウトを30秒に設定(デフォルトは5秒)
                proxy_host = '10.10.0.51',
                proxy_port = 22,
                proxy_username = username,
                proxy_password = password,
                session_log=f"session_logs/{hostname or ip_address}.log",
            ) as ssh:
                # コマンドを実行してJSONデータを取得
                json_output = ssh.send_command('show interfaces extensive | display json | no-more', read_timeout=90) #タイムアウト90秒

        except Exception as e:
            print(f"ホスト {hostname or ip_address} への接続中にエラーが発生しました:", e)
            continue

    # JSONデータからマルチキャストパケットとブロードキャストパケットの値を取り出す関数
    def extract_packet_counts(json_data):
        interface_data = []
        #  interface = json.loads(json_output)['configuration']['interfaces']['interface']
        data = json.loads(json_data)
        input_multicasts = data['interface-information'][0]['physical-interface'][0]['ethernet-mac-statistics'][0]['input-multicasts'][0]['data']
        output_multicasts = data['interface-information'][0]['physical-interface'][0]['ethernet-mac-statistics'][0]['output-multicasts'][0]['data']
        input_broadcasts = data['interface-information'][0]['physical-interface'][0]['ethernet-mac-statistics'][0]['input-broadcasts'][0]['data']
        output_broadcasts = data['interface-information'][0]['physical-interface'][0]['ethernet-mac-statistics'][0]['output-broadcasts'][0]['data']
        input_errors = data['interface-information'][0]['physical-interface'][0]['input-error-list'][0]['input-errors'][0]['data']
        input_drops = data['interface-information'][0]['physical-interface'][0]['input-error-list'][0]['input-drops'][0]['data']
        framing_drops = data['interface-information'][0]['physical-interface'][0]['input-error-list'][0]['framing-errors'][0]['data']
        input_runts = data['interface-information'][0]['physical-interface'][0]['input-error-list'][0]['input-runts'][0]['data']
        input_discards = data['interface-information'][0]['physical-interface'][0]['input-error-list'][0]['input-discards'][0]['data']
        input_l3_incompletes = data['interface-information'][0]['physical-interface'][0]['input-error-list'][0]['input-l3-incompletes'][0]['data']
        input_l2_channel_errors = data['interface-information'][0]['physical-interface'][0]['input-error-list'][0]['input-l2-channel-errors'][0]['data']
        input_l2_mismatch_timeouts = data['interface-information'][0]['physical-interface'][0]['input-error-list'][0]['input-l2-mismatch-timeouts'][0]['data']
        input_fifo_errors = data['interface-information'][0]['physical-interface'][0]['input-error-list'][0]['input-fifo-errors'][0]['data']
        input_resource_errors = data['interface-information'][0]['physical-interface'][0]['input-error-list'][0]['input-resource-errors'][0]['data']
        return input_multicasts, output_multicasts, input_broadcasts, output_broadcasts, input_errors, input_drops, framing_drops, input_runts, input_discards, input_l3_incompletes, input_l2_channel_errors, input_l2_mismatch_timeouts, input_fifo_errors, input_resource_errors

    # JSONデータから値を取得
    input_multicasts, output_multicasts, input_broadcasts, output_broadcasts, input_errors, input_drops, framing_drops, input_runts, input_discards, input_l3_incompletes, input_l2_channel_errors, input_l2_mismatch_timeouts, input_fifo_errors, input_resource_errors = extract_packet_counts(json_output)

    # 結果の出力
    print(f"ホスト: {hostname or ip_address}")
    print("マルチキャストパケット（入力）:", input_multicasts)
    print("マルチキャストパケット（出力）:", output_multicasts)
    print("ブロードキャストパケット（入力）:", input_broadcasts)
    print("ブロードキャストパケット（出力）:", output_broadcasts)
    print(input_errors)
    print(input_drops)
    print(framing_drops)
    print(input_runts)
    print(input_discards)
    print(input_l3_incompletes)
    print(input_l2_channel_errors)
    print(input_l2_mismatch_timeouts)
    print(input_fifo_errors)
    print(input_resource_errors)
    print('-' * 50)
