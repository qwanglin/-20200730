# -*- coding: utf-8 -*-
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from common.WriteToTxt import *
class NewDriver:
    def __init__(self,driver):
        self.driver=driver
    def findElement(self,by,value):
        """

        :param loc:
        :return:
        """
        try:
            element = WebDriverWait(self.driver, int(30)).until(
                EC.presence_of_element_located((by,value))
            )
        except Exception as e:
            logWriteToTxt('找到不到此元素%s 加载页面失败'%(value,))
            raise e
        else:
            # log.logger.info('The page of %s had already find the element %s' % (self, loc))
            return self.driver.find_element(by,value)

    def findElements(self,by,value):
            """

            :param loc:
            :return:
            """
            try:
                element = WebDriverWait(self.driver, int(30)).until(
                    EC.presence_of_element_located((by,value))
                )

            except Exception as e:

                logWriteToTxt('找到不到此元素%s 加载页面失败'%(value,))
                raise e

            else:

                return self.driver.find_elements(by,value)