import unittest
import os
from unittest.mock import patch, MagicMock
import openai
from Generate import Config, BaiduAPI, OpenAIAPI, AIService

class TestConfig(unittest.TestCase):
    """测试配置类"""
    
    def setUp(self):
        """测试前的设置"""
        # 保存原始环境变量
        self.original_env = {}
        for key in ['BAIDU_API_KEY', 'BAIDU_SECRET_KEY', 'OPENAI_API_KEY', 'OPENAI_MODEL', 'OPENAI_TEMPERATURE', 'OPENAI_MAX_TOKENS']:
            self.original_env[key] = os.getenv(key)
            
        # 设置测试环境变量
        os.environ['BAIDU_API_KEY'] = 'test_baidu_key'
        os.environ['BAIDU_SECRET_KEY'] = 'test_baidu_secret'
        os.environ['OPENAI_API_KEY'] = 'test_openai_key'
        os.environ['OPENAI_MODEL'] = 'test-model'
        os.environ['OPENAI_TEMPERATURE'] = '0.5'
        os.environ['OPENAI_MAX_TOKENS'] = '1000'
        os.environ['OPENAI_API_BASE'] = 'https://test.openai.com/v1'
        os.environ['USE_BAIDU'] = 'true'
        
        # 重新加载配置
        Config.reload()

    def tearDown(self):
        """测试后的清理"""
        # 恢复原始环境变量
        os.environ.clear()
        for key, value in self.original_env.items():
            if value is not None:
                os.environ[key] = value
        
        # 重新加载配置
        Config.reload()

    def test_config_initialization(self):
        """测试配置初始化"""
        print("\n开始测试配置初始化...")
        self.assertEqual(Config.BAIDU_API_KEY, 'test_baidu_key')
        self.assertEqual(Config.BAIDU_SECRET_KEY, 'test_baidu_secret')
        self.assertEqual(Config.OPENAI_API_KEY, 'test_openai_key')
        self.assertEqual(Config.OPENAI_API_BASE, 'https://test.openai.com/v1')
        self.assertTrue(Config.USE_BAIDU)
        self.assertEqual(Config.OPENAI_MODEL, 'test-model')
        self.assertEqual(Config.OPENAI_TEMPERATURE, 0.5)
        self.assertEqual(Config.OPENAI_MAX_TOKENS, 1000)
        print("配置初始化测试完成")

class TestBaiduAPI(unittest.TestCase):
    """测试百度API类"""

    @patch('requests.post')
    def test_get_access_token(self, mock_post):
        """测试获取访问令牌"""
        print("\n开始测试获取访问令牌...")
        # 模拟成功响应
        mock_post.return_value.json.return_value = {
            'access_token': 'test_token'
        }
        
        token = BaiduAPI.get_access_token()
        self.assertEqual(token, 'test_token')
        
        # 模拟失败响应
        mock_post.return_value.json.return_value = {
            'error': 'error',
            'error_description': 'test error'
        }
        
        with self.assertRaises(Exception):
            BaiduAPI.get_access_token()
        print("获取访问令牌测试完成")

    @patch('Generate.BaiduAPI.get_access_token')
    def test_call_api(self, mock_get_token):
        """测试API调用"""
        print("\n开始测试百度API调用...")
        # 模拟访问令牌
        mock_get_token.return_value = 'test_token'
        
        # 模拟请求失败
        with patch('requests.post') as mock_post:
            mock_post.side_effect = Exception('test error')
            with self.assertRaises(Exception):
                BaiduAPI.call_api('test prompt')
        print("百度API调用测试完成")

class TestOpenAIAPI(unittest.TestCase):
    """测试OpenAI API类"""

    def test_initialize(self):
        """测试初始化"""
        print("\n开始测试OpenAI API初始化...")
        OpenAIAPI.initialize()
        self.assertEqual(openai.api_key, Config.OPENAI_API_KEY)
        self.assertEqual(openai.api_base, Config.OPENAI_API_BASE)
        print("OpenAI API初始化测试完成")

    @patch('openai.ChatCompletion.create')
    def test_call_api(self, mock_create):
        """测试API调用"""
        print("\n开始测试OpenAI API调用...")
        # 模拟OpenAI API响应
        mock_create.return_value = MagicMock(
            choices=[MagicMock(message=MagicMock(content='test response'))]
        )
        
        response = OpenAIAPI.call_api('test prompt')
        self.assertEqual(response, 'test response')
        print("OpenAI API调用测试完成")

class TestAIService(unittest.TestCase):
    """测试AI服务类"""

    def test_get_ai_provider(self):
        """测试获取AI提供商"""
        print("\n开始测试获取AI提供商...")
        # 测试使用百度API
        Config.USE_BAIDU = True
        provider = AIService.get_ai_provider()
        self.assertEqual(provider, BaiduAPI)

        # 测试使用OpenAI API
        Config.USE_BAIDU = False
        provider = AIService.get_ai_provider()
        self.assertEqual(provider, OpenAIAPI)
        print("获取AI提供商测试完成")

    @patch('Generate.AIService.get_ai_provider')
    def test_generate_solution(self, mock_get_provider):
        """测试生成解决方案"""
        print("\n开始测试生成解决方案...")
        mock_provider = MagicMock()
        mock_provider.call_api.return_value = 'test solution'
        mock_get_provider.return_value = mock_provider

        solution = AIService.generate_solution('test content')
        self.assertEqual(solution, 'test solution')
        print("生成解决方案测试完成")

    @patch('Generate.AIService.get_ai_provider')
    def test_shorten_text(self, mock_get_provider):
        """测试文本缩短"""
        print("\n开始测试文本缩短...")
        mock_provider = MagicMock()
        mock_provider.call_api.return_value = 'short text'
        mock_get_provider.return_value = mock_provider

        result = AIService.shorten_text('long text')
        self.assertEqual(result, 'short text')
        print("文本缩短测试完成")

if __name__ == '__main__':
    unittest.main()
