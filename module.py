import os, sys
import time
from slack_sdk import WebClient  # Slack APIへリクエストするためのクライアント。SDK使用。
from slack_sdk.errors import SlackApiError  # Slack APIエラーオブジェクト。SDK使用。
import shutil

def copy_to_onedrive():
    print(os.path.exists("C:/Users/SSSRC/公立大学法人大阪/SSSRC - 入退室履歴"))
    shutil.copy("./history.xlsx", "C:/Users/SSSRC/公立大学法人大阪/SSSRC - 入退室履歴/history.xlsx")

def delete_old_messages():
  TERM = 60 * 60 * 24 * 4
  channel_id = "C02V60L6ED8"
  client = WebClient(
    token="xoxp-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
  )  # WebClientインスタンス生成。引数は、Tokenコード。
  latest = int(time.time() - TERM)  # 現在日時 - １週間 の UNIX時間
  cursor = None  # シーク位置。最初は None ページを指定して、次からは next_cursor が指し示す位置。
  while True:
    try:
      response = client.conversations_history(  # conversations_history ＝ チャット一覧を得る
        channel=channel_id,
        latest=latest,
        cursor=cursor  # チャンネルID、latest、シーク位置を指定。
        # latestに指定した時間よりも古いメッセージが得られる。latestはUNIX時間で指定する。
      )
    except SlackApiError as e:
      #print("error")
      sys.exit(
        e.response["error"]  # エラーが発生したら即終了
      )  # str like 'invalid_auth', 'channel_not_found'
    #print("messages" in response)
    if "messages" in response:  # response["messages"]が有るか？
      #print(response)
      i = 0
      for message in response["messages"]:  # response["messages"]が有る場合、１件ずつループ
        time.sleep(0.1)
        #print(message)
        try:
          if i == 3:
            break
          #print("delete")
          client.chat_delete(
            channel=channel_id, ts=message["ts"]
          )  # chat_delete ＝ 指定したチャットを削除
          # 引数にチャンネルID、ts（タイムスタンプ：conversations_historyのレスポンスに含まれる）を指定して、削除
          i += 1
        except SlackApiError as e:
          print("error")
          sys.exit(e.response["error"])  # エラーが発生したら即終了

    if "has_more" not in response or response["has_more"] is not True:
      # conversations_historyのレスポンスに["has_more"]が無かったり、has_moreの値がFalseだった場合、終了する。
      break

    if (
      "response_metadata" in response
      and "next_cursor" in response["response_metadata"]
    ):  # conversations_historyのレスポンスに["response_metadata"]["next_cursor"]が有る場合、cursorをセット
      # （上に戻って、もう一度、conversations_history取得）
      cursor = response["response_metadata"]["next_cursor"]
    else:
      break
    time.sleep(0.1)


if __name__ == "__main__":
  delete_old_messages()
  copy_to_onedrive()
