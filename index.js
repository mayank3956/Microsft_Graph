
import axios from 'axios';
import config from '../config/configurations';
import BaseProcessor from '../RedisProcessor/BaseProcessor';
import { getRecords } from '../lib/utilities';
class MicrosoftGraphService  {
  async getImageDataFromMSG(search, option) {
    try {
      async function baseConerter(resultStr) {
        return new Buffer.from(resultStr, 'binary').toString('base64');
      }
      const tokenResponse = await cca.acquireTokenByClientCredential(
        tokenRequest,
      );
      const token = tokenResponse.accessToken;
      const finalResponse = [];
      const firstHalfRecords = ['user1@gmail.com', 'user2@gmail.com', 'user3@gmail.com']
      const firstHalfRequests = [];
      const secondHalfRequest = [];
      const secondHalfRecords = [];
      // Microsoft Batch API can send the maximux 20 request at once time.
     // That's why we dividing the records array in two parts because according to our projects we can not send more than 24 request at once time.
      if (firstHalfRecords.length > 1) {
        while (firstHalfRecords.length > secondHalfRecords.length) {
          secondHalfRecords.push(firstHalfRecords.splice(firstHalfRecords.length - 1)[0]);
        }
      }
      firstHalfRecords.map((userId) => {
        const req = {
          url: `users/${userId}/photo/$value`,
          method: 'GET',
          id: userId,
        };
        firstHalfRequests.push(req);
        return req;
      });
      const firstHalfModifiyObjects = Object.assign({ requests: firstHalfRequests });
      const firstHalfModifiyObjects = Object.assign({ requests: firstHalfRequests });
      const firstHalfResponses = await axios({
        method: 'post',
        url: `https://graph.microsoft.com/v1.0/$batch`,
        data: firstHalfModifiyObjects, 
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });
      firstHalfResponses.data.responses.map((ele) => {
        const images = {
          email: ele.id,
          IMG64BIT: `data:image/png;base64, ${ele.body}`,
        };
        finalResponse.push(images);
      });
      if (firstHalfRecords.length > 1) {
        secondHalfRecords.map((userId) => {
          const req = {
            url: `users/${userId}/photo/$value`,
            method: 'GET',
            id: userId,
          };
          secondHalfRequest.push(req);
          return req;
        });
        const secondHalfModifiyObjects = Object.assign({ requests: secondHalfRequest });
        const SecondHalfResponses =  await axios({
        method: 'post',
        url: `${url}/v1.0/$batch`,
        data: secondHalfModifiyObjects, 
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });;
        SecondHalfResponses.data.responses.map((ele) => {
          // if(ele.IMG64BIT){

          // }
          const images = {
            email: ele.id,
            IMG64BIT: `data:image/png;base64, ${ele.body}`,
          };
          finalResponse.push(images);
        });
      }
      return finalResponse;
    } catch (error) {
      logger.error(
        'MicrosoftGraphService :: getImageDataFromMSG :: error',
        error,
      );
      throw error.response.statusText;
    }
  }
// Get Image of particular user
  async getImageDataFromEmail(email) {
    try {
      const imagesRes = [];
      const tokenResponse = await cca.acquireTokenByClientCredential(
        tokenRequest,
      );
      const token = tokenResponse.accessToken;
      const request = [];
      const req = {
        url: `users/${email}/photo/$value`,
        method: 'GET',
        id: email,
      };
      request.push(req);
      const batchReq = Object.assign({ requests: request });
      const finalRespnse =  await axios({
        method: 'post',
        url: `https://graph.microsoft.com/v1.0/$batch`,
        data: secondHalfModifiyObjects, 
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });;
      finalRespnse.data.responses.map((ele) => {
        const images = {
          email: ele.id,
          IMG64BIT: `data:image/png;base64, ${ele.body}`,
        };
        imagesRes.push(images);
      });
      return imagesRes;
    } catch (error) {
      logger.error(
        'MicrosoftGraphService :: getImageDataFromMSG :: error',
        error.response.data,
      );
      throw error.response.data;
    }
  }
}

export default MicrosoftGraphService;
