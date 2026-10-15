import os
import sys

from dotenv import load_dotenv


# 必须在 import Agent 前加载 .env
load_dotenv()


from agent import PPTAgent


def check_environment():

    api_key = os.getenv(
        "OPENAI_API_KEY"
    )

    if not api_key:

        print(
            "错误：没有找到 "
            "OPENAI_API_KEY"
        )

        print(
            "请在项目目录中的 "
            ".env 文件填写："
        )

        print()

        print(
            "OPENAI_API_KEY="
            "sk-xxxxxxxx"
        )

        return False

    return True


def print_header():

    print()
    print("=" * 60)

    print(
        "PowerPoint to PDF (PPT2PDF) Agent"
    )

    print("=" * 60)

    print(
        "你可以直接输入自然语言。"
    )

    print()

    print(
        "示例："
    )

    print(
        "把 method.pptx 第 5 页"
        "导出成 PDF，裁掉白边"
    )

    print()

    print(
        "把 method.pptx "
        "第 3、5 到 7 页导出来，"
        "所有页面保持一样大小"
    )

    print()

    print(
        "把 D:/paper/test.pptx "
        "全部页面导出，白边保留 3pt"
    )

    print()

    print(
        "命令："
    )

    print(
        "  exit   退出"
    )

    print(
        "  reset  清除对话上下文"
    )

    print("=" * 60)
    print()


def main():

    if not check_environment():
        sys.exit(1)

    agent = PPTAgent(
        verbose=True
    )

    print_header()

    while True:

        try:

            user_input = input(
                "请输入 > "
            ).strip()

        except (
            KeyboardInterrupt,
            EOFError
        ):

            print()
            print("退出。")
            break

        if not user_input:
            continue

        # ------------------------
        # exit
        # ------------------------

        if user_input.lower() in {
            "exit",
            "quit",
            "q"
        }:

            print("退出。")
            break

        # ------------------------
        # reset
        # ------------------------

        if user_input.lower() == "reset":

            agent.reset()

            print(
                "已清除 Agent "
                "对话上下文。"
            )

            continue

        # ------------------------
        # Agent
        # ------------------------

        try:

            answer = agent.run(
                user_input
            )

            print()
            print(
                f"Agent > {answer}"
            )
            print()

        except Exception as e:

            print()
            print(
                "Agent 运行失败："
            )

            print(
                f"{type(e).__name__}: "
                f"{e}"
            )

            print()


if __name__ == "__main__":
    main()
    # check_environment()