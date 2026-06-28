"""Quick smoke test for t2cad_llm + t2cad_pipeline."""
import sys
sys.path.insert(0, r"C:\Users\zhaot\.text_to_cad")

from t2cad_llm import LLMClient, load_config, strip_code_fence

cfg = load_config()
print(f"Provider: {cfg['provider']}, Model: {cfg['model']}")

client = LLMClient(cfg)

# Test 1: Basic chat
print("\n[Test 1] Basic chat...")
try:
    resp = client.chat([{"role": "user", "content": "回复：OK"}])
    print(f"  Response: {resp.strip()[:80]}")
    print("  PASSED")
except Exception as e:
    print(f"  FAILED: {e}")

# Test 2: Memory
print("\n[Test 2] Memory...")
try:
    client.clear_memory("test")
    r1 = client.chat_with_memory("我叫赵天兵", session_id="test")
    print(f"  R1: {r1.strip()[:80]}")
    r2 = client.chat_with_memory("我叫什么名字？", session_id="test")
    print(f"  R2: {r2.strip()[:80]}")
    print("  PASSED" if "赵天兵" in r2 else "  PARTIAL (memory may not be working)")
except Exception as e:
    print(f"  FAILED: {e}")

# Test 3: Pipeline
print("\n[Test 3] CodeGenPipeline...")
from t2cad_pipeline import CodeGenPipeline
pipeline = CodeGenPipeline(client)

# Simulate a simple exec namespace
def hello(name):
    return f"Hello, {name}!"

result = pipeline.run(
    system_prompt="你是Python专家，只输出Python代码。",
    snapshot="",
    user_input="调用 hello('World') 函数",
    exec_namespace={"hello": hello},
    fixer_prompt="修复Python代码错误。只输出修正后的代码。",
    max_retries=2,
)
print(f"  Success: {result['success']}")
print(f"  Result: {result['result']}")
print(f"  Attempt: {result['attempt']}")

print("\n=== All tests complete ===")
