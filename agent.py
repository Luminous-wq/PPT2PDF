import json
import os

from openai import OpenAI

from tools.ppt2pdf import export_ppt_to_cropped_pdf
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PROMPT_PATH = BASE_DIR / "prompt.txt"

SYSTEM_PROMPT = PROMPT_PATH.read_text(encoding="utf-8")

# Responses API function schema
TOOLS = [
    {
        "type": "function",

        "name": "export_ppt_to_cropped_pdf",

        "description": (
            "从 PowerPoint 文件中导出指定页面为 PDF，"
            "并自动裁剪页面四周白色空白。"
            "支持单页、多页、全部页面以及统一裁剪尺寸。"
        ),

        "parameters": {
            "type": "object",

            "properties": {

                "ppt_path": {
                    "type": "string",
                    "description": (
                        "PPT 或 PPTX 文件路径。"
                        "可以是绝对路径或相对路径。"
                    )
                },

                "pages": {
                    "type": "string",
                    "description": (
                        "需要导出的 PPT 页码。"
                        "例如 '3'、'3,5-7'、'all'。"
                    )
                },

                "output_path": {
                    "type": "string",
                    "description": (
                        "最终 PDF 输出路径。"
                        "如果用户没有指定则不要提供。"
                    )
                },

                "margin": {
                    "type": "number",
                    "description": (
                        "裁掉白边以后额外保留的边距，"
                        "单位 pt。默认 6。"
                    )
                },

                "crop_mode": {
                    "type": "string",
                    "enum": [
                        "per_page",
                        "union"
                    ],
                    "description": (
                        "per_page 表示每页独立裁剪；"
                        "union 表示多页采用相同裁剪区域。"
                    )
                }
            },

            "required": [
                "ppt_path"
            ],

            "additionalProperties": False
        }
    }
]


class PPTAgent:

    def __init__(
        self,
        model=None,
        verbose=True
    ):

        # self.client = OpenAI()
        self.client = OpenAI(
            base_url="https://openrouter.ai/api/v1",
            api_key=os.getenv("OPENROUTER_API_KEY")
        )

        self.model = (
            model
            or os.getenv(
                "OPENAI_MODEL",
                "gpt-5.6"
            )
        )

        self.verbose = verbose

        # 保存当前会话上下文
        self.input_items = []

    def reset(self):
        """
        清空 Agent 对话上下文。
        """

        self.input_items = []

    def execute_tool(
        self,
        tool_name,
        arguments
    ):
        
        if tool_name == "export_ppt_to_cropped_pdf":

            return export_ppt_to_cropped_pdf(
                **arguments
            )

        raise ValueError(
            f"未知工具：{tool_name}"
        )

    def _print_tool_call(
        self,
        name,
        arguments
    ):

        if not self.verbose:
            return

        print()
        print("=" * 60)
        print("Agent 调用工具")
        print("=" * 60)

        print(
            f"Tool: {name}"
        )

        print(
            json.dumps(
                arguments,
                ensure_ascii=False,
                indent=2
            )
        )

        print("=" * 60)
        print()

    def _print_tool_result(
        self,
        result
    ):

        if not self.verbose:
            return

        print()
        print("=" * 60)
        print("Tool 执行结果")
        print("=" * 60)

        print(
            json.dumps(
                result,
                ensure_ascii=False,
                indent=2
            )
        )

        print("=" * 60)
        print()

    def run(
        self,
        user_message,
        max_tool_rounds=10
    ):

        # 加入用户输入
        self.input_items.append(
            {
                "role": "user",
                "content": user_message
            }
        )

        for _ in range(
            max_tool_rounds
        ):

            # -------------------------
            # 让模型决定下一步
            # -------------------------

            response = (
                self.client.responses.create(
                    model=self.model,
                    instructions=SYSTEM_PROMPT,
                    tools=TOOLS,
                    input=self.input_items
                )
            )

            # 官方 function calling 流程中，
            # model output 要继续放入后续 input
            self.input_items += (
                response.output
            )

            # 找出工具调用
            function_calls = [item for item in response.output if item.type == "function_call"]

            # -------------------------
            # 没有 Tool Call
            # 说明模型已经给出最终回答
            # -------------------------

            if not function_calls:

                text = (
                    response.output_text
                    or
                    "任务已经完成。"
                )

                return text

            # -------------------------
            # 执行所有 Tool Calls
            # -------------------------

            for call in function_calls:

                try:
                    arguments = json.loads(
                        call.arguments
                    )

                except Exception as e:

                    arguments = {}

                    tool_result = {
                        "success": False,
                        "error": (
                            "工具参数 JSON "
                            f"解析失败：{e}"
                        )
                    }

                else:

                    self._print_tool_call(
                        call.name,
                        arguments
                    )

                    try:

                        result = (
                            self.execute_tool(
                                call.name,
                                arguments
                            )
                        )

                        if isinstance(
                            result,
                            dict
                        ):
                            tool_result = result

                        else:
                            tool_result = {
                                "success": True,
                                "result": result
                            }

                    except Exception as e:

                        tool_result = {
                            "success": False,
                            "tool": call.name,
                            "error_type": (
                                type(e).__name__
                            ),
                            "error": str(e)
                        }

                self._print_tool_result(
                    tool_result
                )

                # -------------------------
                # 把 Python 的执行结果
                # 返回给模型
                # -------------------------

                self.input_items.append(
                    {
                        "type":
                            "function_call_output",

                        "call_id":
                            call.call_id,

                        "output":
                            json.dumps(
                                tool_result,
                                ensure_ascii=False
                            )
                    }
                )

        return (
            "Agent 连续调用工具次数过多，"
            "已停止执行。"
        )